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

from collections.abc import AsyncIterator, Callable, Sequence
from pathlib import PurePosixPath

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
from core.errors import (
    DefaultDriveNotFoundError,
    DriveItemNotFoundError,
    DriveNotFoundError,
    FailedWhenCreateDriveItemError,
    FileAlreadyExistError,
    FileVeryLargeError,
    GraphResponseError,
    InvalidConflictBehaviorError,
    InvalidRemoteNameError,
    LocalFileNotReadableError,
    LocalPathIsDirectoryError,
    LocalPathNotFoundError,
    MSCronosError,
    NotAFileError,
    NotAFolderError,
    SiteResolutionError,
    SmallFileUploadError,
    TreeDirectoryCreationError,
    TreeFileUploadError,
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
)
from core.utils import rename_with_uuid

# Tamanho maximo entregue ao `LargeFileUploadTask` em cada envio parcial.
MAX_CHUNK_SIZE = 60 * 1024 * 1024

MAX_SMALL_FILE_SIZE = 250_000_000  # 250 milhões de bytes


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

            response = await self._client_manager.client.sites.with_url(
                secure_url
            ).get()

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
                response = await self._client_manager.client.sites.by_site_id(
                    site.id
                ).drives.get()

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

            current_page = await page_iterator.next()

            if not current_page:
                break

            page_iterator.current_page = current_page

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
            response = await self._client_manager.client.sites.by_site_id(
                site.id
            ).drive.get()

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
            response = await self._client_manager.client.drives.by_drive_id(
                drive_id
            ).get()

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
            site_response = await self._client_manager.client.drives.by_drive_id(
                library.id
            ).root.get()

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
                response = (
                    await self._client_manager.client.drives.by_drive_id(library.id)
                    .items.by_drive_item_id(parent.id)
                    .children.get()
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

            if len(current_page.value) == 0:
                return

            for item in current_page.value:
                if not isinstance(item, DriveItem):
                    raise TypeError(
                        "A pagina de filhos retornou um item de tipo inesperado: "
                        f"'{type(item).__name__}'."
                    )
                page_content.append(parse_drive_item(item))
            yield SharePointItemCollection.from_collection(page_content)

            current_page = await page_iterator.next()

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
        self, library, parent, name: str
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

            response = (
                await self._client_manager.client.drives.by_drive_id(library.id)
                .items.by_drive_item_id(parent.id)
                .children.post(body)
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

        try:
            content_bytes = local_file.path.read_bytes()
        except OSError as error:
            raise LocalFileNotReadableError(
                f"Nao foi possivel ler o arquivo local '{local_file.path}'."
            ) from error

        try:
            existing_file = await self.find_file_by_name(
                library,
                parent,
                local_file.name,
            )

            remote_name = local_file.name
            should_create = existing_file is None

            if existing_file is not None and conflict_behavior == "fail":
                raise FileAlreadyExistError(
                    f"Ja existe um arquivo chamado '{local_file.name}' sob o "
                    f"item remoto '{parent.id}', e a politica de conflito e 'fail'."
                )
            if existing_file is not None and conflict_behavior == "rename":
                remote_name = rename_with_uuid(local_file.name)
                should_create = True

            if should_create:
                content_fragment = build_create_content_url(remote_name)
                create_url = build_drive_create_content_url(
                    library.id,
                    parent.id,
                    content_fragment,
                )
                response = await (
                    self._client_manager.client.drives.by_drive_id(library.id)
                    .items.by_drive_item_id(parent.id)
                    .content.with_url(create_url)
                    .put(content_bytes)
                )
            else:
                if existing_file is None:
                    raise SmallFileUploadError(
                        "Nao foi possivel substituir o arquivo porque nenhum "
                        f"item remoto corresponde a '{local_file.name}'."
                    )
                response = await (
                    self._client_manager.client.drives.by_drive_id(library.id)
                    .items.by_drive_item_id(existing_file.id)
                    .content.put(content_bytes)
                )
                remote_name = existing_file.name or local_file.name

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
                remote_name=uploaded_item.name or remote_name,
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

        try:
            upload_session = await self._create_upload_session(
                library=library,
                parent=parent,
                local_file=local_file,
                conflict_behavior=conflict_behavior,
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
            else "",
            conflict_behavior=conflict_behavior,
        )

    async def _create_upload_session(
        self,
        library: DocumentLibrary,
        parent: SharePointItem,
        local_file: LocalFile,
        conflict_behavior: ConflictBehavior = "fail",
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

        uploadable_propieties = DriveItemUploadableProperties(
            additional_data={"@microsoft.graph.conflictBehavior": conflict_behavior}
        )

        request_body = CreateUploadSessionPostRequestBody(item=uploadable_propieties)

        upload_session = await (
            self._client_manager.client.drives.by_drive_id(library.id)
            .items.by_drive_item_id(f"{parent.id}:/{local_file.name}:")
            .create_upload_session.post(request_body)
        )

        if upload_session is None:
            raise RuntimeError(
                "O Microsoft Graph nao retornou uma sessao para o upload de "
                f"'{local_file.name}' sob o item '{parent.id}'."
            )

        if upload_session.upload_url is None:
            raise RuntimeError(
                "A sessao de upload retornada pelo Microsoft Graph nao contem "
                f"upload_url para o arquivo '{local_file.name}'."
            )

        return upload_session

    async def upload_tree(
        self,
        parent: SharePointItem,
        library: DocumentLibrary,
        staging_tree: StagingFilesystemTree,
        conflict_behavior: ConflictBehavior = "fail",
    ) -> dict[PurePosixPath, dict[SharePointItem, list[FileUploadResult]]]:
        """Cria a estrutura remota e envia os arquivos nivel por nivel.

        A primeira fase resolve todos os diretorios da arvore de staging. A
        segunda usa o mapa resultante para enviar sequencialmente os arquivos
        de cada nivel. A operacao e fail-fast e nao desfaz itens ja criados ou
        enviados quando uma etapa posterior falha.
        """
        uploaded_tree: dict[
            PurePosixPath, dict[SharePointItem, list[FileUploadResult]]
        ] = {}

        try:
            remote_tree = await self._ensure_tree(
                parent=parent,
                library=library,
                staging_tree=staging_tree,
                conflict_behavior=conflict_behavior,
            )
        except TreeDirectoryCreationError:
            raise
        except MSCronosError as error:
            raise TreeUploadError(
                "Nao foi possivel preparar os diretorios remotos antes do "
                "upload da arvore local "
                f"'{staging_tree.source.root.path}'."
            ) from error

        for level in staging_tree.levels:
            remote_parent = remote_tree.get(level.relative_path)

            if not remote_parent:
                raise TreeDirectoryCreationError(
                    "O nivel de staging "
                    f"'{level.relative_path}' nao possui um diretorio remoto "
                    "resolvido apos a criacao da arvore."
                )

            uploaded_files: list[FileUploadResult] = []
            for file in level.staging_files:
                try:
                    uploaded_file = await self.upload(
                        library=library,
                        parent=remote_parent,
                        local_file=file.source,
                        conflict_behavior=file.conflict_behavior,
                    )
                except MSCronosError as error:
                    raise TreeFileUploadError(
                        f"Falha ao enviar o arquivo local '{file.source.path}' "
                        f"para o nivel remoto '{level.relative_path}' "
                        f"(item pai '{remote_parent.id}')."
                    ) from error
                else:
                    uploaded_files.append(uploaded_file)

            # Niveis sem arquivos tambem permanecem representados no resultado.
            uploaded_tree[level.relative_path] = {remote_parent: uploaded_files}

        return uploaded_tree

    async def _ensure_tree(
        self,
        parent: SharePointItem,
        library: DocumentLibrary,
        staging_tree: StagingFilesystemTree,
        conflict_behavior: ConflictBehavior = "fail",
    ) -> dict[PurePosixPath, SharePointItem]:
        """Garante cada nivel remoto e o associa ao caminho relativo local."""
        remote_tree: dict[PurePosixPath, SharePointItem] = {}
        if not parent.is_folder:
            raise NotAFolderError(
                "Nao foi possivel criar a arvore remota porque o item pai "
                f"'{parent.id}' nao representa uma pasta."
            )

        for level in staging_tree.levels:
            target_path = level.relative_path

            if staging_tree.target_root != PurePosixPath("."):
                target_path = staging_tree.target_root / level.relative_path

            try:
                remote_level = await self.ensure_remote_folder_path(
                    library=library,
                    root=parent,
                    folders_parts=target_path.parts,
                    conflict_behavior=conflict_behavior,
                )
            except MSCronosError as error:
                raise TreeDirectoryCreationError(
                    "Falha ao garantir o diretorio remoto do nivel "
                    f"'{level.relative_path}' no caminho de destino "
                    f"'{target_path}' da biblioteca '{library.id}'."
                ) from error
            except ODataError as error:
                parsed_error = parse_o_data_error(
                    error,
                    operation="ensure_tree_directory",
                )
                raise TreeDirectoryCreationError(
                    "O Microsoft Graph rejeitou a criacao ou resolucao do "
                    f"nivel '{level.relative_path}' no destino '{target_path}'."
                ) from parsed_error

            remote_tree[level.relative_path] = remote_level

        return remote_tree
