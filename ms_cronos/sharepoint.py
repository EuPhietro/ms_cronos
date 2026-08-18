"""Servico de alto nivel para operacoes do SharePoint via Microsoft Graph.

Este modulo adapta o SDK do Graph aos contratos semanticos do Core. O objetivo
e expor operacoes pequenas e previsiveis para navegacao, criacao de pastas e
upload de arquivos pequenos ou grandes.

Exemplo de listagem acumulada:
    service = SharePointService(graph_client_manager)
    site = await service.resolve_site('https://tenant.sharepoint.com/sites/RH')
    drive = await service.get_default_drive(site)
    root = await service.get_drive_root(drive)
    children = await service.list_children(drive, root)

Exemplo de consumo pagina a pagina:
    async for page in service.iter_children(drive, root):
        for item in page:
            print(item.name)
"""

from __future__ import annotations

import asyncio
import hashlib
import inspect
from collections.abc import AsyncIterator, Awaitable, Callable, Sequence
from dataclasses import dataclass, field
from pathlib import Path, PurePosixPath
from typing import TypeVar

import httpx
from msgraph.generated.drives.item.items.item.create_upload_session.create_upload_session_post_request_body import (
    CreateUploadSessionPostRequestBody,
)
from msgraph.generated.models.drive import Drive
from msgraph.generated.models.drive_collection_response import DriveCollectionResponse
from msgraph.generated.models.drive_item import DriveItem
from msgraph.generated.models.drive_item_collection_response import (
    DriveItemCollectionResponse,
)
from msgraph.generated.models.drive_item_uploadable_properties import (
    DriveItemUploadableProperties,
)
from msgraph.generated.models.o_data_errors.o_data_error import ODataError
from msgraph.generated.models.upload_session import UploadSession
from msgraph_core import PageIterator
from msgraph_core.tasks import LargeFileUploadTask

from core.builders import build_folder_drive_item, normalize_conflict_behavior
from core.checkpoint import TreeUploadCheckpointStore
from core.errors import (
    CheckpointMismatchError,
    DefaultDriveNotFoundError,
    DriveItemNotFoundError,
    DriveNotFoundError,
    FailedWhenCreateDriveItemError,
    FileAlreadyExistError,
    FileVeryLargeError,
    GraphResponseError,
    GraphTransportError,
    InvalidConflictBehaviorError,
    InvalidRemoteNameError,
    LocalFileNotReadableError,
    LocalPathIsDirectoryError,
    LocalPathNotFoundError,
    NotAFileError,
    NotAFolderError,
    SiteResolutionError,
    SmallFileUploadError,
    TreeDirectoryCreationError,
    TreeFileUploadError,
    TreeUploadCancelledError,
    TreeUploadError,
    UploadError,
)
from core.graph_client import GraphClientManager
from core.models import (
    ConflictBehavior,
    DocumentLibrary,
    DocumentLibraryCollection,
    FileUploadResult,
    LocalFile,
    SharePointItem,
    SharePointItemCollection,
    SharePointSite,
    StagingFilesystemTree,
    TreeUploadProgress,
    TreeUploadProgressCallback,
    TreeUploadResult,
)
from core.parse import (
    adapt_site,
    parse_drive,
    parse_drive_collection_response,
    parse_drive_item,
    parse_o_data_error,
    parse_site,
)
from core.urls import (
    build_create_content_url,
    build_drive_create_content_url,
    build_graph_site_url,
    validate_remote_path,
)
from core.utils import rename_with_uuid

MAX_CHUNK_SIZE = 10 * 1024 * 1024
MAX_SMALL_FILE_SIZE = 10 * 1024 * 1024
MAX_GRAPH_ATTEMPTS = 4
MAX_RETRY_DELAY_SECONDS = 30.0
RETRYABLE_GRAPH_STATUS_CODES = frozenset({408, 429, 500, 502, 503, 504, 509})

_GraphResultT = TypeVar("_GraphResultT")


@dataclass
class _RemoteTreeState:
    """Estado interno da arvore remota durante uma execucao."""

    result: TreeUploadResult
    children_by_level: dict[PurePosixPath, dict[str, SharePointItem]] = field(
        default_factory=dict
    )


class SharePointService:
    """Orquestra operacoes semanticas do SharePoint sobre o Microsoft Graph.

    O servico funciona como fronteira entre a aplicacao e o SDK. Ele recebe
    referencias internas do Core, executa as chamadas remotas, valida as
    respostas e converte os models do Graph antes de devolve-los ao consumidor.
    Tipos como ``Site``, ``Drive``, ``DriveItem`` e ``PageResult`` permanecem
    restritos a implementacao.

    As operacoes publicas estao organizadas por responsabilidade:

    - resolucao de URLs do SharePoint em ``SharePointSite``;
    - consulta e busca de drives de um site;
    - navegacao pela raiz e pelos filhos de um drive;
    - busca de arquivos e pastas remotas;
    - criacao e garantia de caminhos de pastas;
    - upload de arquivos pequenos ou grandes.

    Listagens paginadas usam o ``PageIterator`` internamente. A classe controla
    a conversao de cada pagina para modelos semanticos e nao entrega o iterador
    do SDK como parte de sua API publica. Erros ``ODataError`` devem ser
    traduzidos para a hierarquia de erros do Core na fronteira de cada operacao.

    Args:
        graph_client_manager: Manager autenticado que fornece o
            ``GraphServiceClient`` e seu ``RequestAdapter``.

    Attributes:
        _client_manager: Manager usado por todas as requisicoes ao Graph.

    Note:
        A instancia nao cria nem fecha o manager recebido. O chamador continua
        responsavel por seu ciclo de vida, preferencialmente com ``async with``.

    Example:
        async with GraphClientManager(credentials) as manager:
            service = SharePointService(manager)
            site = await service.resolve_site(sharepoint_url)
            drive = await service.get_default_drive(site)
            root = await service.get_drive_root(drive)
            children = await service.list_children(drive, root)
    """

    def __init__(self, graph_client_manager: GraphClientManager) -> None:
        """Inicializa o servico com um manager Graph ja configurado.

        Args:
            graph_client_manager: Manager que fornece o cliente autenticado e
                o request adapter usado pelas operacoes do servico.
        """
        self._client_manager = graph_client_manager

    async def _execute_graph_operation(
        self,
        operation: Callable[[], Awaitable[_GraphResultT]],
        *,
        operation_name: str,
        retry_transport: bool = True,
    ) -> _GraphResultT:
        """Executa uma chamada Graph com retry para falhas transitórias."""
        for attempt in range(1, MAX_GRAPH_ATTEMPTS + 1):
            try:
                return await operation()
            except ODataError as error:
                if (
                    error.response_status_code not in RETRYABLE_GRAPH_STATUS_CODES
                    or attempt == MAX_GRAPH_ATTEMPTS
                ):
                    raise
                delay = self._retry_delay(error.response_headers, attempt)
            except httpx.TransportError as error:
                if not retry_transport or attempt == MAX_GRAPH_ATTEMPTS:
                    raise GraphTransportError(
                        f"Falha de transporte durante '{operation_name}' apos "
                        f"{attempt} tentativa(s): {error}."
                    ) from error
                delay = self._retry_delay(None, attempt)

            await asyncio.sleep(delay)

        raise GraphTransportError(
            f"A operacao '{operation_name}' excedeu o limite de tentativas."
        )

    @staticmethod
    def _retry_delay(
        headers: dict[str, str] | None,
        attempt: int,
    ) -> float:
        """Prioriza `Retry-After` e usa backoff exponencial como fallback."""
        if headers:
            retry_after = next(
                (
                    value
                    for key, value in headers.items()
                    if key.casefold() == "retry-after"
                ),
                None,
            )
            if retry_after is not None:
                try:
                    return min(float(retry_after), MAX_RETRY_DELAY_SECONDS)
                except ValueError:
                    pass
        return min(float(2 ** (attempt - 1)), MAX_RETRY_DELAY_SECONDS)

    @staticmethod
    def _bind_tree_result(
        result: TreeUploadResult,
        *,
        parent: SharePointItem,
        library: DocumentLibrary,
        staging_tree: StagingFilesystemTree,
    ) -> None:
        """Vincula um checkpoint a uma unica origem e destino remoto."""
        expected = {
            "source_root": staging_tree.source.root.path.resolve(),
            "library_id": library.id,
            "parent_item_id": parent.id,
            "target_root": staging_tree.target_root,
            "staging_fingerprint": SharePointService._staging_fingerprint(staging_tree),
        }

        for attribute, expected_value in expected.items():
            current_value = getattr(result, attribute)
            if current_value is not None and current_value != expected_value:
                raise CheckpointMismatchError(
                    "O checkpoint nao pertence a esta operacao de upload: "
                    f"'{attribute}' esperava {expected_value!r}, mas contem "
                    f"{current_value!r}."
                )
            setattr(result, attribute, expected_value)

    @staticmethod
    def _staging_fingerprint(staging_tree: StagingFilesystemTree) -> str:
        """Resume estrutura, nomes, tamanhos e conflitos sem ler os arquivos."""
        entries: list[str] = []
        for level in staging_tree.levels:
            entries.append(f"level:{level.relative_path.as_posix()}")
            entries.extend(
                f"folder:{folder.relative_path.as_posix()}:{folder.remote_name}"
                for folder in level.staging_folders
            )
            entries.extend(
                "file:"
                f"{file.relative_path.as_posix()}:{file.remote_name}:"
                f"{file.source.size}:{file.conflict_behavior}"
                for file in level.staging_files
            )

        digest = hashlib.sha256()
        for entry in sorted(entries):
            digest.update(entry.encode("utf-8"))
            digest.update(b"\0")
        return digest.hexdigest()

    @staticmethod
    def _raise_if_tree_cancelled(
        cancel_event: asyncio.Event | None,
        result: TreeUploadResult,
    ) -> None:
        """Interrompe o fluxo em um ponto seguro preservando o progresso."""
        if cancel_event is not None and cancel_event.is_set():
            raise TreeUploadCancelledError(
                "O upload da arvore foi cancelado pelo chamador.",
                partial_result=result,
            )

    @staticmethod
    async def _notify_tree_progress(
        callback: TreeUploadProgressCallback | None,
        progress: TreeUploadProgress,
    ) -> None:
        """Aceita callbacks sincrononos ou assincronos de progresso."""
        if callback is None:
            return
        callback_result = callback(progress)
        if inspect.isawaitable(callback_result):
            await callback_result

    # Resolucao de sites

    async def resolve_site(
        self,
        sharepoint_url: str,
    ) -> SharePointSite:
        """Resolve uma URL humana do SharePoint e devolve um `SharePointSite`.

        A rota `sites.with_url(...)` pode retornar dados em formatos diferentes
        dependendo do SDK; por isso este metodo aceita tanto `response.value`
        quanto `response.additional_data`.

        Args:
            sharepoint_url: URL humana do site no SharePoint.

        Returns:
            Referencia semantica do site resolvido.

        Raises:
            SiteResolutionError: Se o Graph nao devolver uma resposta.
            GraphResponseError: Se a resposta nao contiver dados suficientes.
            MSCronosError: Se uma falha OData for traduzida pelo Core.
        """
        try:
            secure_url = build_graph_site_url(sharepoint_url)

            response = await self._execute_graph_operation(
                lambda: self._client_manager.client.sites.with_url(secure_url).get(),
                operation_name=f"resolver site '{sharepoint_url}'",
            )

            if not response:
                raise SiteResolutionError(
                    "O Microsoft Graph nao retornou resposta ao resolver o "
                    f"site '{sharepoint_url}'."
                )

            # O SDK pode preencher `value` ou somente `additional_data` nesta rota.
            if response.value:
                data = parse_site(response.value[0])
            elif response.additional_data.get("id"):
                data = adapt_site(response.additional_data)
            else:
                raise GraphResponseError(
                    "A resposta de resolucao do site nao contem dados para "
                    f"montar SharePointSite: '{sharepoint_url}'."
                )

        except SiteResolutionError:
            raise
        except GraphResponseError:
            raise
        except ODataError as error:
            raise parse_o_data_error(error, operation="resolve_site")
        else:
            return data

    # Consulta de drives

    async def list_site_drives(
        self,
        site: SharePointSite,
        *,
        pagination: bool = True,
        max_pages: int | None = None,
    ) -> DocumentLibraryCollection:
        """Lista os drives associados a um site, com paginacao opcional.

        Quando ``pagination`` for ``False``, somente a primeira resposta do
        Graph sera convertida. Quando for ``True``, o metodo
        percorrera os links de continuacao ate a ultima pagina ou ate atingir
        ``max_pages``. O limite conta a primeira pagina.

        Args:
            site: Referencia semantica do site que possui os drives.
            pagination: Controla se as paginas seguintes devem ser consultadas.
            max_pages: Quantidade maxima de paginas processadas no cliente.

        Returns:
            Colecao semantica contendo os drives das paginas processadas.

        Raises:
            ValueError: Se ``max_pages`` for menor ou igual a zero.
            DriveNotFoundError: Se o Graph nao devolver o envelope inicial.
            GraphResponseError: Se alguma pagina nao puder ser interpretada.
            MSCronosError: Se uma falha OData for traduzida pelo Core.
        """

        pages: list[DocumentLibrary] = []

        async def get_first_page(
            site: SharePointSite,
        ) -> DriveCollectionResponse:
            """Solicita e valida o envelope inicial devolvido pelo Graph."""
            try:
                response = await self._execute_graph_operation(
                    lambda: self._client_manager.client.sites.by_site_id(
                        site.id
                    ).drives.get(),
                    operation_name=f"listar bibliotecas do site '{site.id}'",
                )

                if not response:
                    raise DriveNotFoundError(
                        "O Microsoft Graph nao retornou resposta ao listar as "
                        f"bibliotecas do site '{site.id}'."
                    )

                if not response.value:
                    raise GraphResponseError(
                        "A primeira pagina de bibliotecas do site "
                        f"'{site.id}' nao contem itens no campo 'value'."
                    )

            except DriveNotFoundError:
                raise
            except ODataError as error:
                raise parse_o_data_error(error, operation="list_site_drives")
            else:
                return response

        if max_pages is not None and max_pages <= 0:
            raise ValueError(
                f"max_pages deve ser maior que zero; recebido: {max_pages}."
            )

        first_page = await get_first_page(site=site)

        if not pagination:
            return parse_drive_collection_response(first_page)

        page_iterator = PageIterator(
            request_adapter=self._client_manager.client.request_adapter,
            response=first_page,
        )

        current_page = page_iterator.current_page
        page_counter = 0

        while current_page is not None:
            if not current_page.value:
                raise GraphResponseError(
                    "Uma pagina da listagem de bibliotecas nao contem itens "
                    "no campo 'value'."
                )

            for item in current_page.value:
                if not isinstance(item, Drive):
                    raise GraphResponseError(
                        "A listagem de bibliotecas retornou um item de tipo "
                        f"inesperado: '{type(item).__name__}'."
                    )

                pages.append(parse_drive(item))

            page_counter += 1

            if max_pages is not None and page_counter >= max_pages:
                break

            current_page = await self._execute_graph_operation(
                page_iterator.next,
                operation_name=f"obter proxima pagina de bibliotecas de '{site.id}'",
            )

            if not current_page:
                break

        return DocumentLibraryCollection.from_collection(pages)

    async def find_drive_by_name(
        self,
        site: SharePointSite,
        name: str,
    ) -> DocumentLibrary | None:
        """Busca um drive pelo nome entre as bibliotecas de um site.

        A busca solicita a listagem paginada completa do site e compara
        ``DocumentLibrary.name`` com o nome informado. A comparacao atual e exata e
        diferencia letras maiusculas de minusculas.

        Args:
            name: Nome exato do drive que deve ser localizado.
            site: Referencia semantica do site que sera pesquisado.

        Returns:
            O primeiro ``DocumentLibrary`` com nome correspondente ou ``None`` quando
            nenhum drive satisfizer a busca.

        Raises:
            DriveNotFoundError: Se a listagem de drives do site falhar por
                ausencia do recurso.
            MSCronosError: Se a camada de listagem traduzir uma falha do Graph
                para um erro semantico do Core.
        """
        try:
            drivers = await self.list_site_drives(
                site=site,
                pagination=True,
            )

            for drive in drivers:
                if drive.name == name:
                    return drive

        except DriveNotFoundError:
            raise
        except ODataError as error:
            raise parse_o_data_error(
                o_data_error=error,
                operation="get_drive_by_name",
            )
        except Exception:
            raise

    async def get_default_drive(
        self,
        site: SharePointSite,
    ) -> DocumentLibrary:
        """Obtem o drive padrao de um site resolvido.

        Args:
            site: Site cuja biblioteca padrao sera consultada.

        Returns:
            Referencia semantica do drive padrao.

        Raises:
            DefaultDriveNotFoundError: Se o Graph nao devolver o drive.
            MSCronosError: Se uma falha OData for traduzida pelo Core.
        """
        try:
            response = await self._execute_graph_operation(
                lambda: self._client_manager.client.sites.by_site_id(
                    site.id
                ).drive.get(),
                operation_name=f"obter biblioteca padrao do site '{site.id}'",
            )

            if not response:
                raise DefaultDriveNotFoundError(
                    "O Microsoft Graph nao retornou a biblioteca padrao do "
                    f"site '{site.id}'."
                )

        except ODataError as error:
            raise parse_o_data_error(error, operation="get_default_drive")
        else:
            return parse_drive(response)

    async def get_drive_by_id(
        self,
        drive_id: str,
    ) -> DocumentLibrary:
        """Consulta um drive especifico pelo identificador.

        Use quando voce ja tem o id de um drive e quer confirmar ou atualizar os
        metadados basicos desse drive.

        Args:
            drive_id: Identificador do drive que sera consultado.

        Returns:
            Nova referencia semantica com os metadados retornados pelo Graph.

        Raises:
            DriveNotFoundError: Se o drive nao for encontrado.
            MSCronosError: Se uma falha OData for traduzida pelo Core.
        """
        try:
            response = await self._execute_graph_operation(
                lambda: self._client_manager.client.drives.by_drive_id(drive_id).get(),
                operation_name=f"obter biblioteca '{drive_id}'",
            )

            if not response:
                raise DriveNotFoundError(
                    f"O Microsoft Graph nao retornou a biblioteca com id '{drive_id}'."
                )

        except DriveNotFoundError:
            raise
        except ODataError as error:
            raise parse_o_data_error(error, operation="get_drive_by_id")
        else:
            return parse_drive(response)

    # Navegacao em itens do drive

    async def get_drive_root(
        self,
        library: DocumentLibrary,
    ) -> SharePointItem:
        """Obtem o item que representa a raiz de um drive.

        Args:
            library: Drive cuja raiz sera consultada.

        Returns:
            Referencia semantica do ``DriveItem`` raiz.

        Raises:
            DriveNotFoundError: Se o Graph nao devolver a raiz.
            MSCronosError: Se uma falha OData for traduzida pelo Core.
        """
        try:
            site_response = await self._execute_graph_operation(
                lambda: self._client_manager.client.drives.by_drive_id(
                    library.id
                ).root.get(),
                operation_name=f"obter raiz da biblioteca '{library.id}'",
            )

            if not site_response:
                raise DriveNotFoundError(
                    "O Microsoft Graph nao retornou o item raiz da biblioteca "
                    f"'{library.id}'."
                )

        except DriveNotFoundError:
            raise
        except ODataError as error:
            raise parse_o_data_error(error, operation="get_drive_root")
        else:
            return parse_drive_item(site_response)

    async def iter_children(
        self,
        library: DocumentLibrary,
        parent: SharePointItem,
    ) -> AsyncIterator[SharePointItemCollection]:
        """Percorre, sob demanda, as paginas de filhos de um item remoto.

        O metodo encapsula o ``PageIterator`` do SDK e converte cada pagina em
        uma ``SharePointItemCollection``. A proxima requisicao somente acontece
        quando o consumidor avanca o ``async for``.

        Args:
            library: Biblioteca que contem o item remoto consultado.
            parent: Item cuja relacao ``children`` sera percorrida.

        Yields:
            Uma colecao semantica para cada pagina nao vazia devolvida pelo
            Microsoft Graph.

        Raises:
            GraphResponseError: Se o Graph nao devolver o envelope inicial.
            TypeError: Se uma pagina contiver um objeto diferente de
                ``DriveItem``.
            MSCronosError: Se uma falha OData for traduzida pelo Core.

        Note:
            Uma primeira resposta sem itens encerra o gerador sem produzir
            colecoes. Este metodo nao acumula paginas em memoria.
        """

        async def get_first_page(
            library: DocumentLibrary, parent: SharePointItem
        ) -> DriveItemCollectionResponse | None:
            """Solicita o envelope inicial usado para criar o paginador."""

            try:
                response = await self._execute_graph_operation(
                    lambda: (
                        self._client_manager.client.drives.by_drive_id(library.id)
                        .items.by_drive_item_id(parent.id)
                        .children.get()
                    ),
                    operation_name=(
                        f"listar filhos do item '{parent.id}' na biblioteca "
                        f"'{library.id}'"
                    ),
                )

                if not response:
                    raise GraphResponseError(
                        "O Microsoft Graph nao retornou resposta ao listar os "
                        f"filhos do item '{parent.id}' na biblioteca "
                        f"'{library.id}'."
                    )

                if not response.value:
                    return None

            except ODataError as error:
                raise parse_o_data_error(error, operation="iter_children")
            else:
                return response

        response = await get_first_page(library=library, parent=parent)

        if response is None:
            return

        page_iterator = PageIterator(
            response=response,
            request_adapter=self._client_manager.client.request_adapter,
        )

        current_page = page_iterator.current_page

        while current_page is not None:
            page_content = []
            if current_page.value is None:
                break

            if current_page.value:
                for item in current_page.value:
                    if not isinstance(item, DriveItem):
                        raise TypeError(
                            "A pagina de filhos retornou um item de tipo inesperado: "
                            f"'{type(item).__name__}'."
                        )
                    page_content.append(parse_drive_item(item))
                yield SharePointItemCollection.from_collection(page_content)

            current_page = await self._execute_graph_operation(
                page_iterator.next,
                operation_name=(
                    f"obter proxima pagina de filhos do item '{parent.id}'"
                ),
            )

    async def list_children(
        self,
        library: DocumentLibrary,
        parent: SharePointItem,
        filter: Callable[[SharePointItem], bool] | None = None,
    ) -> SharePointItemCollection:
        """Acumula os filhos imediatos de um item remoto.

        O metodo consome todas as paginas produzidas por ``iter_children`` e
        materializa uma unica ``SharePointItemCollection``. O filtro opcional e
        aplicado a cada pagina antes da acumulacao. A operacao nao e recursiva.

        Args:
            library: Biblioteca que contem o item consultado.
            parent: Item remoto cuja relacao ``children`` sera acumulada.
            filter: Predicado opcional. Somente itens para os quais ele retornar
                ``True`` serao incluidos.

        Returns:
            Colecao com os itens aceitos de todas as paginas. Quando nao houver
            filhos ou correspondencias, retorna uma colecao vazia.

        Raises:
            GraphResponseError: Se o Graph nao devolver o envelope inicial.
            MSCronosError: Se uma falha OData for traduzida pelo Core.
        """

        pages = []
        try:
            async for page in self.iter_children(library=library, parent=parent):
                if not filter:
                    pages.extend(page)
                    continue
                pages.extend(item for item in page if filter(item))

        except Exception:
            raise

        return SharePointItemCollection.from_collection(pages)

    # Listagens especializadas de filhos

    async def get_children_folder(
        self,
        library: DocumentLibrary,
        parent: SharePointItem,
    ) -> SharePointItemCollection:
        """Lista somente as pastas filhas imediatas de um item remoto.

        Todas as paginas sao consumidas por ``list_children`` e filtradas pela
        facet semantica ``is_folder``.

        Args:
            library: Biblioteca que contem o item consultado.
            parent: Item remoto cuja colecao de filhos sera filtrada.

        Returns:
            Colecao contendo apenas os filhos classificados como pasta.

        Raises:
            GraphResponseError: Se a listagem nao devolver um envelope valido.
            MSCronosError: Se uma falha OData for traduzida pelo Core.
        """
        try:
            childrens = await self.list_children(
                library=library, parent=parent, filter=lambda item: item.is_folder
            )
            return childrens

        except GraphResponseError:
            raise
        except ODataError as error:
            raise parse_o_data_error(error, operation="get_children_folder")

    async def get_children_file(
        self,
        library: DocumentLibrary,
        parent: SharePointItem,
    ) -> SharePointItemCollection:
        """Lista somente os arquivos filhos imediatos de um item remoto.

        Todas as paginas sao consumidas por ``list_children`` e filtradas pela
        facet semantica ``is_file``.

        Args:
            library: Biblioteca que contem o item consultado.
            parent: Item remoto cuja colecao de filhos sera filtrada.

        Returns:
            Colecao contendo apenas os filhos classificados como arquivo.

        Raises:
            GraphResponseError: Se a listagem nao devolver um envelope valido.
            MSCronosError: Se uma falha OData for traduzida pelo Core.
        """
        try:
            children = await self.list_children(
                library=library, parent=parent, filter=lambda item: item.is_file
            )
            return children
        except GraphResponseError:
            raise
        except ODataError as error:
            raise parse_o_data_error(error, operation="get_children_file")

    # Busca de filhos imediatos

    async def find_child_by_name(
        self, library: DocumentLibrary, parent: SharePointItem, name: str
    ) -> SharePointItem | None:
        """Busca um filho imediato pelo nome dentro de uma pasta remota.

        Args:
            library: Drive que contem o item pai.
            parent: Pasta em que a busca sera executada.
            name: Nome exato do filho procurado.

        Returns:
            Primeiro filho correspondente ou ``None`` quando nao encontrado.

        Raises:
            GraphResponseError: Se a listagem de filhos falhar.
            MSCronosError: Se uma falha OData for traduzida pelo Core.
        """
        try:
            children = await self.list_children(
                library, parent, filter=lambda item: item.name == name
            )

            for child in children:
                if child.name == name:
                    return child

            return None

        except NotAFolderError:
            raise
        except DriveItemNotFoundError:
            raise
        except ODataError as error:
            raise parse_o_data_error(error, operation="list_children")

    async def find_child_by_id(
        self,
        library: DocumentLibrary,
        parent: SharePointItem,
        item_id: str,
    ) -> SharePointItem | None:
        """Busca um filho imediato pelo identificador do DriveItem.

        Args:
            library: Drive que contem o item pai.
            parent: Pasta em que a busca sera executada.
            item_id: Identificador exato do filho procurado.

        Returns:
            Filho correspondente ou ``None`` quando nao encontrado.

        Raises:
            GraphResponseError: Se a listagem de filhos falhar.
            MSCronosError: Se uma falha OData for traduzida pelo Core.
        """
        try:
            children = await self.list_children(
                library, parent, filter=lambda item: item.id == item_id
            )

            for child in children:
                if child.id == item_id:
                    return child

            return None
        except NotAFolderError:
            raise
        except DriveItemNotFoundError:
            raise
        except ODataError as error:
            raise parse_o_data_error(error, operation="list_children")

    async def find_child_by_web_url(
        self,
        library: DocumentLibrary,
        parent: SharePointItem,
        web_url: str,
    ) -> SharePointItem | None:
        """Busca um filho imediato por sua URL web.

        Args:
            library: Drive que contem o item pai.
            parent: Pasta em que a busca sera executada.
            web_url: URL web exata do filho procurado.

        Returns:
            Filho correspondente ou ``None`` quando nao encontrado.

        Raises:
            GraphResponseError: Se a listagem de filhos falhar.
            MSCronosError: Se uma falha OData for traduzida pelo Core.
        """
        try:
            children = await self.list_children(
                library, parent, filter=lambda item: item.web_url == web_url
            )

            for child in children:
                if child.web_url == web_url:
                    return child

            return None
        except NotAFolderError:
            raise
        except DriveItemNotFoundError:
            raise
        except ODataError as error:
            raise parse_o_data_error(error, operation="list_children")

    async def find_folder_by_name(
        self,
        library: DocumentLibrary,
        parent: SharePointItem,
        name: str,
    ) -> SharePointItem | None:
        """Localiza a primeira pasta filha com o nome informado.

        Args:
            library: Biblioteca que contem o item consultado.
            parent: Item remoto cujas pastas filhas serao pesquisadas.
            name: Nome exato da pasta procurada.

        Returns:
            A primeira pasta correspondente ou ``None`` quando nao encontrada.

        Raises:
            GraphResponseError: Se a listagem de filhos falhar.
            MSCronosError: Se uma falha OData for traduzida pelo Core.
        """

        try:
            folders = await self.get_children_folder(library=library, parent=parent)
            candidates = [folder for folder in folders if folder.name == name]
            if len(candidates) > 0:
                return candidates[0]
            return
        except GraphResponseError:
            raise
        except ODataError as error:
            raise parse_o_data_error(
                error,
                operation="find_folder_by_name",
            ) from error

    async def find_file_by_name(
        self,
        library: DocumentLibrary,
        parent: SharePointItem,
        name: str,
    ) -> SharePointItem | None:
        """Localiza o primeiro arquivo filho com o nome informado.

        Args:
            library: Biblioteca que contem o item consultado.
            parent: Item remoto cujos arquivos filhos serao pesquisados.
            name: Nome exato do arquivo procurado.

        Returns:
            O primeiro arquivo correspondente ou ``None`` quando nao encontrado.

        Raises:
            GraphResponseError: Se a listagem de filhos falhar.
            MSCronosError: Se uma falha OData for traduzida pelo Core.
        """

        try:
            files = await self.get_children_file(library=library, parent=parent)
            candidates = [file for file in files if file.name == name]
            if len(candidates) > 0:
                return candidates[0]
            return
        except GraphResponseError:
            raise
        except ODataError as error:
            raise parse_o_data_error(
                error,
                operation="find_file_by_name",
            ) from error

    async def find_child_folder_by_id(
        self,
        library: DocumentLibrary,
        parent: SharePointItem,
        item_id: str,
    ) -> SharePointItem | None:
        """Busca uma pasta filha imediata pelo identificador do DriveItem.

        Args:
            library: Drive que contem a pasta pai.
            parent: Pasta em que a busca sera executada.
            item_id: Identificador exato da pasta procurada.

        Returns:
            Referencia da pasta ou ``None`` quando nao encontrada.

        Raises:
            NotAFolderError: Se o pai nao for pasta ou se o id pertencer a
                um arquivo.
            DriveItemNotFoundError: Se a listagem do pai falhar.
            MSCronosError: Se uma falha OData for traduzida pelo Core.
        """
        try:
            children = await self.list_children(library, parent)

            for child in children:
                if child.id == item_id and child.is_file:
                    raise NotAFolderError(
                        "Conflito de tipo ao buscar pasta filha pelo id "
                        f"'{item_id}': o item encontrado e um arquivo."
                    )
                if child.id == item_id:
                    return child

            return None

        except NotAFolderError:
            raise

        except DriveItemNotFoundError:
            raise
        except ODataError as error:
            raise parse_o_data_error(error, operation="find_child_folder")

    async def find_child_folder_web_url(
        self,
        library: DocumentLibrary,
        parent: SharePointItem,
        web_url: str,
    ) -> SharePointItem | None:
        """Busca uma pasta filha imediata pela URL web do DriveItem.

        Args:
            library: Drive que contem a pasta pai.
            parent: Pasta em que a busca sera executada.
            web_url: URL web exata da pasta procurada.

        Returns:
            Referencia da pasta ou ``None`` quando nao encontrada.

        Raises:
            NotAFolderError: Se o pai nao for pasta ou se a URL pertencer a
                um arquivo.
            DriveItemNotFoundError: Se a listagem do pai falhar.
            MSCronosError: Se uma falha OData for traduzida pelo Core.
        """
        try:
            children = await self.list_children(library, parent)

            for child in children:
                if child.web_url == web_url and child.is_file:
                    raise NotAFolderError(
                        "Conflito de tipo ao buscar pasta filha pela URL "
                        f"'{web_url}': o item encontrado e um arquivo."
                    )
                if child.web_url == web_url:
                    return child

            return None

        except NotAFolderError:
            raise

        except DriveItemNotFoundError:
            raise
        except ODataError as error:
            raise parse_o_data_error(error, operation="find_child_folder")

    # Criacao e garantia de pastas

    async def create_folder(
        self,
        library: DocumentLibrary,
        parent: SharePointItem,
        folder_name: str,
        conflict_behavior: ConflictBehavior = "fail",
    ) -> SharePointItem:
        """Cria uma pasta filha imediata dentro de uma pasta remota.

        Args:
            library: Drive em que a pasta sera criada.
            parent: Pasta remota que recebera o novo filho.
            folder_name: Nome da pasta que sera criada.
            conflict_behavior: Politica aplicada quando o nome ja existir.

        Returns:
            Referencia semantica da pasta criada pelo Graph.

        Raises:
            NotAFolderError: Se o item pai nao representar uma pasta.
            InvalidRemoteNameError: Se o nome remoto for invalido.
            InvalidConflictBehaviorError: Se a politica nao for suportada.
            FailedWhenCreateDriveItemError: Se o Graph nao devolver o item.
            MSCronosError: Se uma falha OData for traduzida pelo Core.
        """
        try:
            if not parent.is_folder:
                raise NotAFolderError(
                    f"O item pai '{parent.id}' nao representa uma pasta remota."
                )

            body = build_folder_drive_item(folder_name, conflict_behavior)

            response = await self._execute_graph_operation(
                lambda: (
                    self._client_manager.client.drives.by_drive_id(library.id)
                    .items.by_drive_item_id(parent.id)
                    .children.post(body)
                ),
                operation_name=(
                    f"criar pasta '{folder_name}' sob o item '{parent.id}'"
                ),
                retry_transport=False,
            )

            if not response:
                raise FailedWhenCreateDriveItemError(
                    "O Microsoft Graph nao retornou o item da pasta "
                    f"'{folder_name}' criada sob o pai '{parent.id}'."
                )
        except (
            NotAFolderError,
            InvalidRemoteNameError,
            InvalidConflictBehaviorError,
            FailedWhenCreateDriveItemError,
        ):
            raise
        except ODataError as error:
            raise parse_o_data_error(error, operation="create_folder")

        else:
            return parse_drive_item(response)

    async def ensure_remote_folder_path(
        self,
        library: DocumentLibrary,
        root: SharePointItem,
        folders_parts: Sequence[str],
        conflict_behavior: ConflictBehavior = "fail",
    ) -> SharePointItem:
        """Garante uma cadeia de pastas remotas a partir de uma pasta raiz.

        Cada parte e buscada entre os filhos imediatos da pasta atual. Pastas
        ausentes sao criadas antes que o metodo avance para o proximo nivel.
        Uma sequencia vazia devolve o proprio ``root``.

        Args:
            library: Drive que contem a hierarquia remota.
            root: Pasta usada como ponto inicial da navegacao.
            folders_parts: Nomes ordenados dos niveis que devem existir.
            conflict_behavior: Politica usada ao criar pastas ausentes.

        Returns:
            Referencia da ultima pasta da cadeia garantida.

        Raises:
            NotAFolderError: Se a raiz ou algum item encontrado nao for pasta.
            DriveItemNotFoundError: Se uma busca intermediaria falhar.
            InvalidRemoteNameError: Se alguma parte tiver nome invalido.
            InvalidConflictBehaviorError: Se a politica nao for suportada.
            FailedWhenCreateDriveItemError: Se uma criacao nao for concluida.
        """
        try:
            if not root.is_folder:
                raise NotAFolderError(
                    f"O item raiz '{root.id}' nao representa uma pasta remota."
                )

            current_part: SharePointItem = root

            for folder_part in folders_parts:
                finded_item = await self.find_folder_by_name(
                    library, current_part, folder_part
                )

                if not finded_item:
                    created_part: SharePointItem = await self.create_folder(
                        library,
                        current_part,
                        folder_part,
                        conflict_behavior,
                    )
                    current_part = created_part
                    continue

                current_part = finded_item

            return current_part
        except (
            NotAFolderError,
            DriveItemNotFoundError,
            InvalidRemoteNameError,
            InvalidConflictBehaviorError,
            FailedWhenCreateDriveItemError,
        ):
            raise

    # Upload de arquivos

    async def upload(
        self,
        library: DocumentLibrary,
        parent: SharePointItem,
        local_file: LocalFile,
        conflict_behavior: ConflictBehavior = "fail",
    ) -> FileUploadResult:
        """Envia um arquivo usando o fluxo adequado ao seu tamanho.

        Arquivos de ate ``MAX_SMALL_FILE_SIZE`` usam upload direto. Arquivos
        maiores usam uma sessao resumivel e sao transmitidos em partes.

        Args:
            library: Drive remoto que recebera o arquivo.
            parent: Pasta remota de destino.
            local_file: Arquivo local preparado para envio.
            conflict_behavior: Politica para nomes remotos ja existentes.

        Returns:
            Resultado semantico com o item remoto e os dados do envio.

        Raises:
            NotAFolderError: Se o destino remoto nao representar uma pasta.
            LocalPathNotFoundError: Se o caminho local nao existir.
            LocalPathIsDirectoryError: Se o caminho local for um diretorio.
            LocalFileNotReadableError: Se o arquivo nao puder ser lido.
            UploadError: Se o fluxo escolhido nao concluir o envio.
        """

        if not parent.is_folder:
            raise NotAFolderError(
                f"O item de destino '{parent.id}' nao representa uma pasta remota."
            )
        if not local_file.path.exists():
            raise LocalPathNotFoundError(
                f"O arquivo local que seria enviado nao existe: '{local_file.path}'."
            )

        if not local_file.path.is_file():
            raise LocalPathIsDirectoryError(
                "O caminho local deve representar um arquivo regular: "
                f"'{local_file.path}'."
            )

        if not local_file.size:
            raise LocalFileNotReadableError(
                f"O arquivo local esta vazio e nao pode ser enviado: "
                f"'{local_file.path}'."
            )

        if local_file.size > MAX_SMALL_FILE_SIZE:
            return await self._upload_large_file(
                library=library,
                parent=parent,
                local_file=local_file,
                conflict_behavior=conflict_behavior,
            )
        return await self._upload_small_file(
            library=library,
            parent=parent,
            local_file=local_file,
            conflict_behavior=conflict_behavior,
        )

    # Implementacao interna de upload

    async def _upload_small_file(
        self,
        library: DocumentLibrary,
        parent: SharePointItem,
        local_file: LocalFile,
        conflict_behavior: ConflictBehavior = "fail",
        *,
        existing_file: SharePointItem | None = None,
        remote_state_known: bool = False,
        remote_name: str | None = None,
    ) -> FileUploadResult:
        """Envia um arquivo pequeno por PUT direto no endpoint `/content`.

        O fluxo decide entre criar ou substituir de acordo com a existencia do
        arquivo remoto e com o `conflict_behavior` informado.

        Args:
            library: Drive remoto que recebera o arquivo.
            parent: Pasta remota de destino.
            local_file: Arquivo local que sera lido integralmente em memoria.
            conflict_behavior: Politica para conflito de nome remoto.

        Returns:
            Resultado semantico do arquivo criado ou substituido.

        Raises:
            NotAFolderError: Se o destino remoto nao representar uma pasta.
            LocalPathNotFoundError: Se o caminho local nao existir.
            LocalPathIsDirectoryError: Se o caminho local for um diretorio.
            LocalFileNotReadableError: Se a leitura local falhar.
            FileVeryLargeError: Se o arquivo exceder o limite deste fluxo.
            FileAlreadyExistError: Se houver conflito no modo ``fail``.
            SmallFileUploadError: Se o Graph nao devolver o item enviado.
            MSCronosError: Se uma falha OData for traduzida pelo Core.
        """
        if not parent.is_folder:
            raise NotAFolderError(
                f"O item de destino '{parent.id}' nao representa uma pasta remota."
            )

        if not local_file.path.exists():
            raise LocalPathNotFoundError(
                f"O caminho local informado nao existe: {local_file.path}"
            )
        if not local_file.path.is_file():
            raise LocalPathIsDirectoryError(
                "O caminho local deve representar um arquivo regular: "
                f"'{local_file.path}'."
            )
        if local_file.size is not None and local_file.size > MAX_SMALL_FILE_SIZE:
            raise FileVeryLargeError(
                f"O arquivo '{local_file.path}' possui {local_file.size} bytes "
                f"e excede o limite de {MAX_SMALL_FILE_SIZE} bytes do upload "
                "pequeno."
            )

        conflict_behavior = normalize_conflict_behavior(conflict_behavior)
        requested_remote_name = remote_name or local_file.name
        build_create_content_url(requested_remote_name)

        try:
            content_bytes = local_file.path.read_bytes()
        except OSError as error:
            raise LocalFileNotReadableError(
                f"Nao foi possivel ler o arquivo local '{local_file.path}'."
            ) from error

        try:
            if not remote_state_known:
                existing_file = await self.find_file_by_name(
                    library,
                    parent,
                    requested_remote_name,
                )

            uploaded_remote_name = requested_remote_name
            should_create = existing_file is None

            if existing_file is not None and conflict_behavior == "fail":
                raise FileAlreadyExistError(
                    f"Ja existe um arquivo chamado '{requested_remote_name}' sob o "
                    f"item remoto '{parent.id}', e a politica de conflito e 'fail'."
                )
            if existing_file is not None and conflict_behavior == "rename":
                uploaded_remote_name = rename_with_uuid(requested_remote_name)
                should_create = True

            if should_create:
                content_fragment = build_create_content_url(uploaded_remote_name)
                create_url = build_drive_create_content_url(
                    library.id,
                    parent.id,
                    content_fragment,
                )
                response = await self._execute_graph_operation(
                    lambda: (
                        self._client_manager.client.drives.by_drive_id(library.id)
                        .items.by_drive_item_id(parent.id)
                        .content.with_url(create_url)
                        .put(content_bytes)
                    ),
                    operation_name=(
                        f"enviar arquivo pequeno '{uploaded_remote_name}' para "
                        f"'{parent.id}'"
                    ),
                )
            else:
                if existing_file is None:
                    raise SmallFileUploadError(
                        "Nao foi possivel substituir o arquivo porque nenhum "
                        f"item remoto corresponde a '{requested_remote_name}'."
                    )
                response = await self._execute_graph_operation(
                    lambda: (
                        self._client_manager.client.drives.by_drive_id(library.id)
                        .items.by_drive_item_id(existing_file.id)
                        .content.put(content_bytes)
                    ),
                    operation_name=(
                        f"substituir conteudo do item '{existing_file.id}'"
                    ),
                )
                uploaded_remote_name = existing_file.name or requested_remote_name

            if not response:
                raise SmallFileUploadError(
                    "O Microsoft Graph nao retornou o item criado ou atualizado "
                    f"apos enviar '{local_file.path}'."
                )

            uploaded_item = parse_drive_item(response)
            return FileUploadResult(
                item=uploaded_item,
                source_path=local_file.path,
                conflict_behavior=conflict_behavior,
                remote_name=uploaded_item.name or uploaded_remote_name,
            )
        except (
            NotAFolderError,
            NotAFileError,
            DriveItemNotFoundError,
            InvalidRemoteNameError,
            InvalidConflictBehaviorError,
            FileAlreadyExistError,
            FileVeryLargeError,
            LocalFileNotReadableError,
            LocalPathIsDirectoryError,
            LocalPathNotFoundError,
            SmallFileUploadError,
        ):
            raise
        except ODataError as error:
            raise parse_o_data_error(
                error,
                operation="upload_small_file",
            ) from error

    async def _upload_large_file(
        self,
        library: DocumentLibrary,
        parent: SharePointItem,
        local_file: LocalFile,
        conflict_behavior: ConflictBehavior = "fail",
        *,
        remote_name: str | None = None,
    ) -> FileUploadResult:
        """Envia um arquivo grande usando uma upload session do Graph.

        O metodo valida o destino e o caminho local, solicita uma sessao
        remota, abre o arquivo como stream binario e delega a divisao e o
        envio dos chunks ao `LargeFileUploadTask`.

        Args:
            library: Drive que contem a pasta remota de destino.
            parent: Pasta remota que recebera o arquivo.
            local_file: Arquivo local preparado para envio.
            conflict_behavior: Comportamento solicitado quando o nome remoto ja
                existir.

        Returns:
            Resultado semantico com a referencia do item remoto enviado.

        Raises:
            NotAFolderError: Se o item remoto de destino nao for uma pasta.
            LocalPathNotFoundError: Se o caminho local nao existir.
            LocalPathIsDirectoryError: Se o caminho local for um diretorio.
            UploadError: Se o Graph nao devolver o item concluido.
            RuntimeError: Se a task nao concluir o upload.
        """
        if not parent.is_folder:
            raise NotAFolderError(
                f"O item de destino '{parent.id}' nao representa uma pasta remota."
            )

        if not local_file.path.exists():
            raise LocalPathNotFoundError(
                f"O arquivo local que seria enviado nao existe: '{local_file.path}'."
            )

        if not local_file.path.is_file():
            raise LocalPathIsDirectoryError(
                "O caminho local deve representar um arquivo regular: "
                f"'{local_file.path}'."
            )

        conflict_behavior = normalize_conflict_behavior(conflict_behavior)
        requested_remote_name = remote_name or local_file.name
        build_create_content_url(requested_remote_name)

        try:
            upload_session = await self._create_upload_session(
                library=library,
                parent=parent,
                local_file=local_file,
                conflict_behavior=conflict_behavior,
                remote_name=requested_remote_name,
            )

            with local_file.path.open("rb") as file_stream:
                upload_task = LargeFileUploadTask(
                    upload_session=upload_session,
                    request_adapter=(self._client_manager.client.request_adapter),
                    stream=file_stream,  # type: ignore
                    parsable_factory=DriveItem,
                    max_chunk_size=MAX_CHUNK_SIZE,
                )
                upload_result = await upload_task.upload()

            if not upload_result.upload_succeeded:
                raise RuntimeError(
                    "A tarefa de upload em partes nao concluiu o envio de "
                    f"'{local_file.path}'."
                )
        except LocalPathNotFoundError:
            raise
        except LocalPathIsDirectoryError:
            raise
        except RuntimeError:
            raise
        except ODataError as error:
            raise parse_o_data_error(error, operation="upload_large_file") from error
        except httpx.TransportError as error:
            raise GraphTransportError(
                f"A conexao foi interrompida durante o upload de '{local_file.path}'."
            ) from error
        if not upload_result.item_response:
            raise UploadError(
                "O upload em partes foi encerrado sem retornar o item remoto "
                f"de '{local_file.path}'."
            )

        uploaded_drive_item = parse_drive_item(upload_result.item_response)

        return FileUploadResult(
            item=uploaded_drive_item,
            source_path=local_file.path,
            remote_name=uploaded_drive_item.name
            if uploaded_drive_item.name is not None
            else requested_remote_name,
            conflict_behavior=conflict_behavior,
        )

    async def _create_upload_session(
        self,
        library: DocumentLibrary,
        parent: SharePointItem,
        local_file: LocalFile,
        conflict_behavior: ConflictBehavior = "fail",
        *,
        remote_name: str | None = None,
    ) -> UploadSession:
        """Cria a sessao remota usada pelo fluxo de upload grande.

        Este helper nao envia bytes. Ele valida o arquivo local, monta o body
        esperado pelo SDK e solicita ao Graph uma `UploadSession` com URL
        pre-autenticada.

        Args:
            library: Drive que contem a pasta remota.
            parent: Pasta remota em que o arquivo sera criado.
            local_file: Arquivo local associado a sessao.
            conflict_behavior: Politica de conflito enviada ao Graph.

        Returns:
            Uma `UploadSession` com `upload_url` preenchida.

        Raises:
            LocalPathNotFoundError: Se o caminho local nao existir.
            LocalPathIsDirectoryError: Se o caminho local for um diretorio.
            LocalFileNotReadableError: Se o arquivo estiver vazio.
            RuntimeError: Se o Graph nao devolver uma sessao utilizavel.
        """
        if not local_file.path.exists():
            raise LocalPathNotFoundError(
                f"Nao e possivel criar uma sessao de upload para um caminho "
                f"inexistente: '{local_file.path}'."
            )

        if not local_file.path.is_file():
            raise LocalPathIsDirectoryError(
                "A sessao de upload exige um arquivo regular, mas recebeu: "
                f"'{local_file.path}'."
            )

        if local_file.size == 0:
            raise LocalFileNotReadableError(
                f"Nao e possivel criar uma sessao para o arquivo vazio "
                f"'{local_file.path}'."
            )

        requested_remote_name = remote_name or local_file.name
        build_create_content_url(requested_remote_name)

        uploadable_propieties = DriveItemUploadableProperties(
            additional_data={"@microsoft.graph.conflictBehavior": conflict_behavior}
        )

        request_body = CreateUploadSessionPostRequestBody(item=uploadable_propieties)

        upload_session = await self._execute_graph_operation(
            lambda: (
                self._client_manager.client.drives.by_drive_id(library.id)
                .items.by_drive_item_id(f"{parent.id}:/{requested_remote_name}:")
                .create_upload_session.post(request_body)
            ),
            operation_name=(
                f"criar sessao de upload para '{requested_remote_name}' sob "
                f"'{parent.id}'"
            ),
        )

        if upload_session is None:
            raise RuntimeError(
                "O Microsoft Graph nao retornou uma sessao para o upload de "
                f"'{requested_remote_name}' sob o item '{parent.id}'."
            )

        if upload_session.upload_url is None:
            raise RuntimeError(
                "A sessao de upload retornada pelo Microsoft Graph nao contem "
                f"upload_url para o arquivo '{requested_remote_name}'."
            )

        return upload_session

    async def upload_tree(
        self,
        parent: SharePointItem,
        library: DocumentLibrary,
        staging_tree: StagingFilesystemTree,
        conflict_behavior: ConflictBehavior = "fail",
        *,
        checkpoint: TreeUploadResult | None = None,
        checkpoint_path: str | Path | None = None,
        checkpoint_interval: int = 100,
        progress_callback: TreeUploadProgressCallback | None = None,
        cancel_event: asyncio.Event | None = None,
    ) -> TreeUploadResult:
        """Cria a estrutura remota e envia os arquivos nivel por nivel.

        A primeira fase resolve todos os diretorios da arvore de staging. A
        segunda reaproveita o indice de filhos de cada nivel para enviar os
        arquivos sem repetir listagens. O progresso pode permanecer somente em
        memoria ou ser persistido periodicamente em um checkpoint JSON.
        """
        if checkpoint_interval <= 0:
            raise ValueError(
                "checkpoint_interval deve ser maior que zero; recebido: "
                f"{checkpoint_interval}."
            )

        checkpoint_store = (
            TreeUploadCheckpointStore(checkpoint_path)
            if checkpoint_path is not None
            else None
        )
        if checkpoint is not None:
            result = checkpoint
        elif checkpoint_store is not None and checkpoint_store.exists:
            result = checkpoint_store.load()
        else:
            result = TreeUploadResult()

        self._bind_tree_result(
            result,
            parent=parent,
            library=library,
            staging_tree=staging_tree,
        )
        total_files = sum(len(level.staging_files) for level in staging_tree.levels)
        total_levels = len(staging_tree.levels)

        await self._notify_tree_progress(
            progress_callback,
            TreeUploadProgress(
                phase="preparing_directories",
                completed_files=result.total_uploaded_files,
                total_files=total_files,
                completed_levels=len(result.completed_levels),
                total_levels=total_levels,
            ),
        )
        try:
            self._raise_if_tree_cancelled(cancel_event, result)
        except TreeUploadCancelledError:
            if checkpoint_store is not None:
                checkpoint_store.save(result)
            raise

        try:
            remote_state = await self._ensure_tree(
                parent=parent,
                library=library,
                staging_tree=staging_tree,
                conflict_behavior=conflict_behavior,
                result=result,
                checkpoint_store=checkpoint_store,
                checkpoint_interval=checkpoint_interval,
                cancel_event=cancel_event,
            )
        except TreeUploadCancelledError:
            if checkpoint_store is not None:
                checkpoint_store.save(result)
            raise
        except TreeDirectoryCreationError as error:
            if error.partial_result is None:
                error.partial_result = result
            if checkpoint_store is not None:
                checkpoint_store.save(result)
            raise
        except Exception as error:
            if checkpoint_store is not None:
                checkpoint_store.save(result)
            raise TreeUploadError(
                "Nao foi possivel preparar os diretorios remotos antes do "
                "upload da arvore local "
                f"'{staging_tree.source.root.path}'.",
                partial_result=result,
            ) from error

        if checkpoint_store is not None:
            checkpoint_store.save(result)

        await self._notify_tree_progress(
            progress_callback,
            TreeUploadProgress(
                phase="uploading_files",
                completed_files=result.total_uploaded_files,
                total_files=total_files,
                completed_levels=len(result.completed_levels),
                total_levels=total_levels,
            ),
        )

        changes_since_checkpoint = 0
        for level in staging_tree.levels:
            try:
                self._raise_if_tree_cancelled(cancel_event, result)
            except TreeUploadCancelledError:
                if checkpoint_store is not None:
                    checkpoint_store.save(result)
                raise

            if level.relative_path in result.completed_levels:
                continue

            remote_parent = result.remote_directories.get(level.relative_path)

            if not remote_parent:
                raise TreeDirectoryCreationError(
                    "O nivel de staging "
                    f"'{level.relative_path}' nao possui um diretorio remoto "
                    "resolvido apos a criacao da arvore.",
                    partial_result=result,
                )

            remote_children = remote_state.children_by_level[level.relative_path]
            uploaded_files = result.uploaded_files.setdefault(
                level.relative_path,
                [],
            )
            completed_sources = {
                upload_result.source_path for upload_result in uploaded_files
            }

            for file in level.staging_files:
                try:
                    self._raise_if_tree_cancelled(cancel_event, result)
                except TreeUploadCancelledError:
                    if checkpoint_store is not None:
                        checkpoint_store.save(result)
                    raise

                if file.source.path in completed_sources:
                    continue

                target_path = staging_tree.target_root / file.relative_path
                validate_remote_path(target_path)
                existing_file = remote_children.get(file.remote_name.casefold())

                try:
                    if file.source.size > MAX_SMALL_FILE_SIZE:
                        uploaded_file = await self._upload_large_file(
                            library=library,
                            parent=remote_parent,
                            local_file=file.source,
                            conflict_behavior=file.conflict_behavior,
                            remote_name=file.remote_name,
                        )
                    else:
                        uploaded_file = await self._upload_small_file(
                            library=library,
                            parent=remote_parent,
                            local_file=file.source,
                            conflict_behavior=file.conflict_behavior,
                            existing_file=existing_file,
                            remote_state_known=True,
                            remote_name=file.remote_name,
                        )
                except Exception as error:
                    if checkpoint_store is not None:
                        checkpoint_store.save(result)
                    raise TreeFileUploadError(
                        f"Falha ao enviar o arquivo local '{file.source.path}' "
                        f"para o nivel remoto '{level.relative_path}' "
                        f"(item pai '{remote_parent.id}').",
                        partial_result=result,
                    ) from error
                else:
                    uploaded_files.append(uploaded_file)
                    remote_children[uploaded_file.remote_name.casefold()] = (
                        uploaded_file.item
                    )
                    changes_since_checkpoint += 1
                    if (
                        checkpoint_store is not None
                        and changes_since_checkpoint >= checkpoint_interval
                    ):
                        checkpoint_store.save(result)
                        changes_since_checkpoint = 0

                    await self._notify_tree_progress(
                        progress_callback,
                        TreeUploadProgress(
                            phase="uploading_files",
                            completed_files=result.total_uploaded_files,
                            total_files=total_files,
                            completed_levels=len(result.completed_levels),
                            total_levels=total_levels,
                            current_path=file.relative_path,
                        ),
                    )

            result.completed_levels.add(level.relative_path)
            changes_since_checkpoint += 1
            if (
                checkpoint_store is not None
                and changes_since_checkpoint >= checkpoint_interval
            ):
                checkpoint_store.save(result)
                changes_since_checkpoint = 0

        if checkpoint_store is not None:
            checkpoint_store.save(result)
        await self._notify_tree_progress(
            progress_callback,
            TreeUploadProgress(
                phase="completed",
                completed_files=result.total_uploaded_files,
                total_files=total_files,
                completed_levels=len(result.completed_levels),
                total_levels=total_levels,
            ),
        )
        return result

    async def _ensure_tree(
        self,
        parent: SharePointItem,
        library: DocumentLibrary,
        staging_tree: StagingFilesystemTree,
        conflict_behavior: ConflictBehavior = "fail",
        *,
        result: TreeUploadResult | None = None,
        checkpoint_store: TreeUploadCheckpointStore | None = None,
        checkpoint_interval: int = 100,
        cancel_event: asyncio.Event | None = None,
    ) -> _RemoteTreeState:
        """Materializa cada pasta uma vez e indexa os filhos de cada nivel."""
        result = result or TreeUploadResult()
        state = _RemoteTreeState(result=result)

        if not parent.is_folder:
            raise NotAFolderError(
                "Nao foi possivel criar a arvore remota porque o item pai "
                f"'{parent.id}' nao representa uma pasta."
            )

        root_relative_path = PurePosixPath(".")
        directories_since_checkpoint = 0
        remote_root = result.remote_directories.get(root_relative_path)
        if remote_root is None:
            try:
                self._raise_if_tree_cancelled(cancel_event, result)
                validate_remote_path(staging_tree.target_root)
                remote_root = await self.ensure_remote_folder_path(
                    library=library,
                    root=parent,
                    folders_parts=staging_tree.target_root.parts,
                    conflict_behavior=conflict_behavior,
                )
            except TreeUploadCancelledError:
                raise
            except Exception as error:
                raise TreeDirectoryCreationError(
                    "Falha ao resolver a raiz remota da arvore no destino "
                    f"'{staging_tree.target_root}'.",
                    partial_result=result,
                ) from error
            result.remote_directories[root_relative_path] = remote_root
            directories_since_checkpoint += 1

        for level in staging_tree.levels:
            self._raise_if_tree_cancelled(cancel_event, result)
            remote_level = result.remote_directories.get(level.relative_path)
            if remote_level is None:
                raise TreeDirectoryCreationError(
                    "A ordem top-down da arvore esta inconsistente: o nivel "
                    f"'{level.relative_path}' foi processado antes de seu pai.",
                    partial_result=result,
                )

            try:
                children = await self.list_children(
                    library=library,
                    parent=remote_level,
                )
                children_index = {
                    child.name.casefold(): child
                    for child in children
                    if child.name is not None
                }
                state.children_by_level[level.relative_path] = children_index

                for folder in level.staging_folders:
                    self._raise_if_tree_cancelled(cancel_event, result)
                    target_path = staging_tree.target_root / folder.relative_path
                    validate_remote_path(target_path)
                    remote_folder = children_index.get(folder.remote_name.casefold())

                    # Um checkpoint ja confirmou esta pasta em uma execucao
                    # anterior. Reaproveita a referencia mesmo se a listagem
                    # atual ainda nao a apresentar por consistencia eventual.
                    if remote_folder is None:
                        remote_folder = result.remote_directories.get(
                            folder.relative_path
                        )

                    if remote_folder is not None and not remote_folder.is_folder:
                        raise NotAFolderError(
                            f"O destino '{target_path}' ja existe como arquivo."
                        )
                    if remote_folder is None:
                        remote_folder = await self.create_folder(
                            library=library,
                            parent=remote_level,
                            folder_name=folder.remote_name,
                            conflict_behavior=conflict_behavior,
                        )
                        children_index[folder.remote_name.casefold()] = remote_folder

                    result.remote_directories[folder.relative_path] = remote_folder
                    directories_since_checkpoint += 1
                    if (
                        checkpoint_store is not None
                        and directories_since_checkpoint >= checkpoint_interval
                    ):
                        checkpoint_store.save(result)
                        directories_since_checkpoint = 0
            except TreeUploadCancelledError:
                raise
            except TreeDirectoryCreationError:
                raise
            except Exception as error:
                raise TreeDirectoryCreationError(
                    "Falha ao materializar as pastas filhas do nivel "
                    f"'{level.relative_path}' na biblioteca '{library.id}'.",
                    partial_result=result,
                ) from error

        return state
