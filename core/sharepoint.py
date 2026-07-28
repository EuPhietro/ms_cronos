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

from core.builders import build_folder_drive_item, build_upload_content
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
    NotAFileError,
    NotAFolderError,
    SiteResolutionError,
    SmallFileUploadError,
    UploadError,
)
from core.graph_client import GraphClientManager
from core.models import (
    ConflictBehavior,
    DocumentLibrary,
    DocumentLibraryCollection,
    FileUploadResult,
    LocalFile,
    PreparedUpload,
    SharePointItem,
    SharePointItemCollection,
    SharePointSite,
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
        # O manager encapsula autenticacao e ciclo de vida do client Graph,
        # deixando o servico focado apenas em regras de SharePoint.
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
            # A URL humana e convertida para a rota Graph que entende
            # `hostname:/server-relative-path`.
            secure_url = build_graph_site_url(sharepoint_url)

            # A rota `sites.with_url(...)` resolve o site sem exigir que o
            # chamador conheca o `site_id` antes.
            response = await self._client_manager.client.sites.with_url(
                secure_url
            ).get()

            # Sem envelope, o Core nao consegue distinguir site inexistente de
            # resposta remota inconsistente.
            if not response:
                raise SiteResolutionError(
                    "Nao foi possivel obter resposta do Graph ao resolver "
                    f"o site: {sharepoint_url}"
                )

            # 4. A resolucao de site e um caso especial do SDK: em alguns
            # cenarios o site resolvido vem em `response.value[0]`; em outros,
            # os dados minimos chegam em `response.additional_data`.
            if response.value:
                # Quando o SDK preenche `value`, a adaptacao segue o parser
                # normal de `Site`.
                data = parse_site(response.value[0])
            elif response.additional_data.get("id"):
                # Alguns cenarios populam apenas `additional_data`; neste caso,
                # o Core faz um fallback controlado para preservar a resolucao.
                data = adapt_site(response.additional_data)
            else:
                raise GraphResponseError(
                    "O Graph respondeu a resolucao do site, mas nao retornou "
                    "dados suficientes para montar um SharePointSite: "
                    f"{sharepoint_url}"
                )

        except SiteResolutionError:
            # Erros internos ja foram montados com contexto suficiente no ponto
            # em que a falha foi detectada.
            raise
        except GraphResponseError:
            # Respostas inconsistentes do Graph tambem permanecem como erro do
            # Core, sem recriar a excecao e perder o traceback original.
            raise
        except ODataError as error:
            # O parser centraliza a traducao do erro remoto para a hierarquia
            # interna do Core.
            raise parse_o_data_error(error, operation="resolve_site")
        else:
            # O retorno final expoe apenas a referencia semantica do Core.
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

        # Acumula apenas modelos internos. Os objetos `Drive` do SDK sao
        # convertidos assim que cada pagina e processada.
        pages: list[DocumentLibrary] = []

        async def get_first_page(
            site: SharePointSite,
        ) -> DriveCollectionResponse:
            """Solicita e valida o envelope inicial devolvido pelo Graph."""
            try:
                # A primeira requisicao parte do id do site. Sua resposta
                # concreta determina o tipo usado pelo `PageIterator` nas
                # requisicoes seguintes.
                response = await self._client_manager.client.sites.by_site_id(
                    site.id
                ).drives.get()

                # Sem envelope nao existe informacao suficiente para iniciar a
                # conversao ou descobrir um eventual `odata_next_link`.
                if not response:
                    raise DriveNotFoundError(
                        "O Graph nao retornou resposta ao listar os drives do "
                        f"site {site.id}."
                    )

                # A primeira pagina deve trazer os itens em `value`. O link de
                # continuacao, quando existir, fica em `odata_next_link`.
                if not response.value:
                    raise GraphResponseError(
                        f"O Graph respondeu a listagem de drives do site {
                            site.id
                        }, mas o envelope veio sem itens em `value`."
                    )

            except DriveNotFoundError:
                # Preserva o erro semantico e seu traceback original.
                raise
            except ODataError as error:
                # Traduz a falha remota antes que um tipo do SDK atravesse a
                # fronteira publica do servico.
                raise parse_o_data_error(error, operation="list_site_drives")
            else:
                # O envelope permanece cru somente dentro deste metodo, pois
                # ainda sera usado para construir o paginador.
                return response

        # Evita um limite ambiguo: zero ou valores negativos nao representam
        # uma quantidade valida de paginas a processar.
        if max_pages is not None and max_pages <= 0:
            raise ValueError(
                "Quando informado, max_pages não pode ser menor ou igual a zero"
            )

        # Toda operacao, paginada ou nao, precisa obter a primeira pagina.
        first_page = await get_first_page(site=site)

        # O fluxo sem paginacao converte somente o primeiro envelope e termina
        # sem consultar seu `odata_next_link`.
        if not pagination:
            return parse_drive_collection_response(first_page)

        # `PageIterator` normaliza o envelope inicial em `PageResult` e usa o
        # request adapter autenticado para buscar os links seguintes.
        page_iterator = PageIterator(
            request_adapter=self._client_manager.client.request_adapter,
            response=first_page,
        )

        # A pagina atual inicial corresponde ao proprio `first_page`; por isso
        # ela deve ser processada antes da primeira chamada a `next()`.
        current_page = page_iterator.current_page

        # O contador representa paginas processadas, e nao quantidade de drives.
        page_counter = 0

        while current_page is not None:
            # Cada `PageResult` guarda os itens crus da pagina em `value`.
            if not current_page.value:
                raise GraphResponseError(
                    "Graph não retornou uma resposa válida para a página atual"
                )

            # Valida e converte cada model do SDK antes de adiciona-lo ao
            # acumulador semantico.
            for item in current_page.value:
                if not isinstance(item, Drive):
                    raise GraphResponseError(
                        f"Erro ao efetuar o Parse de {type(item)} para DocumentLibrary"
                    )

                pages.append(parse_drive(item))

            # A primeira pagina tambem participa do limite informado.
            page_counter += 1

            # Interrompe antes de fazer outra requisicao quando o limite local
            # ja foi satisfeito.
            if max_pages is not None and page_counter >= max_pages:
                break

            # `next()` devolve a proxima pagina ou `None` quando nao existe
            # `odata_next_link`. Na versao instalada, ele nao atualiza sozinho
            # `page_iterator.current_page`; esse estado precisa ser sincronizado
            # antes de solicitar uma terceira pagina.
            current_page = await page_iterator.next()

            if not current_page:
                break

            page_iterator.current_page = current_page

        # Nenhum envelope ou model do SDK e exposto para o consumidor.
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
            # Reutiliza a operacao publica de listagem para manter a paginacao
            # e a conversao dos models do SDK concentradas em um unico metodo.
            drivers = await self.list_site_drives(
                site=site,
                pagination=True,
            )

            # A colecao ja contem `DocumentLibrary`; nenhuma nova conversao do SDK e
            # necessaria durante a busca.
            for drive in drivers:
                # O primeiro nome exatamente igual encerra a busca.
                if drive.name == name:
                    return drive

            # Ao esgotar a colecao, o retorno implicito e `None`, conforme o
            # contrato opcional da assinatura.
        except DriveNotFoundError:
            # Preserva o erro semantico produzido pela operacao de listagem.
            raise
        except ODataError as e:
            # Mantem a traducao centralizada no parser de erros do Core.
            parse_o_data_error(o_data_error=e, operation="get_drive_by_name")
        except Exception:
            # Falhas inesperadas conservam tipo, mensagem e traceback.
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
            # 1. Usa o `site.id` para consultar diretamente o drive padrao
            # do site, sem listar todas as bibliotecas antes.
            response = await self._client_manager.client.sites.by_site_id(
                site.id
            ).drive.get()

            # 2. Nesta rota o Graph devolve um `Drive` unico, nao um envelope
            # de colecao com `value`.
            if not response:
                raise DefaultDriveNotFoundError(
                    f"Drive padrao nao retornado para o site {site.id}."
                )

        except ODataError as error:
            # O parser centraliza a traducao para o erro semantico correto do
            # Core.
            raise parse_o_data_error(error, operation="get_default_drive")
        else:
            # 3. O `Drive` cru retornado pelo SDK e convertido para o contrato
            #    interno enxuto antes de sair do servico.
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
            # 1. Reconsulta diretamente o drive pelo identificador informado.
            response = await self._client_manager.client.drives.by_drive_id(
                drive_id
            ).get()

            # 2. Assim como em `get_default_drive`, aqui a resposta esperada e
            #    um `Drive` unico.
            if not response:
                raise DriveNotFoundError(
                    "O Graph nao retornou o drive solicitado para o "
                    f"identificador {drive_id}."
                )

        except DriveNotFoundError:
            # Se a ausencia foi detectada pelo proprio Core, relancamos a mesma
            # excecao para nao apagar a origem do erro.
            raise
        except ODataError as error:
            # O parser centraliza a traducao para o erro semantico correto do
            # Core.
            raise parse_o_data_error(error, operation="get_drive_by_id")
        else:
            # 3. Converte o `Drive` cru do SDK para `DocumentLibrary`.
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
            # 1. Consulta o item raiz do drive, que no Graph e modelado como um
            #    `DriveItem`.
            site_response = await self._client_manager.client.drives.by_drive_id(
                library.id
            ).root.get()

            # 2. Nesta rota a resposta ja e um item unico, nao um envelope de
            #    colecao.
            if not site_response:
                raise DriveNotFoundError(
                    f"O Graph nao retornou a raiz do drive {library.id}."
                )

        except DriveNotFoundError:
            # A raiz do drive e obrigatoria para navegacao. Se nao veio
            # resposta, o erro interno ja descreve o drive afetado.
            raise
        except ODataError as error:
            # O parser centraliza a traducao para o erro semantico correto do
            # Core.
            raise parse_o_data_error(error, operation="get_drive_root")
        else:
            # 3. Converte o `DriveItem` raiz em `SharePointItem`.
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
                # A consulta parte da biblioteca e do item remoto informados.
                # O envelope concreto inicializa o `PageIterator`.
                response = (
                    await self._client_manager.client.drives.by_drive_id(library.id)
                    .items.by_drive_item_id(parent.id)
                    .children.get()
                )

                # Sem envelope nao existe informacao suficiente para iniciar a
                # conversao ou descobrir um eventual `odata_next_link`.
                if not response:
                    raise GraphResponseError(
                        "O Graph nao retornou resposta ao listar os drives do "
                        f"site {library.id}."
                    )

                # Uma colecao vazia representa ausencia de filhos e encerra o
                # gerador sem produzir pagina semantica.
                if not response.value:
                    return None

            except ODataError as error:
                # Traduz a falha remota antes que um tipo do SDK atravesse a
                # fronteira publica do servico.
                raise parse_o_data_error(error, operation="iter_children")
            else:
                # O envelope permanece cru somente dentro deste metodo, pois
                # ainda sera usado para construir o paginador.

                return response

        response = await get_first_page(library=library, parent=parent)

        # Sem itens na resposta inicial, o gerador termina naturalmente.
        if response is None:
            return

        # O request adapter autenticado permite seguir os links de continuacao.
        page_iterator = PageIterator(
            response=response,
            request_adapter=self._client_manager.client.request_adapter,
        )

        current_page = page_iterator.current_page

        while current_page is not None:
            # Cada pagina e convertida isoladamente para manter o consumo lazy.
            page_content = []
            if current_page.value is None:
                break

            # Uma pagina vazia encerra o fluxo atual sem emitir colecao.
            if len(current_page.value) == 0:
                return

            # O parser unitario reduz cada `DriveItem` ao model interno.
            for item in current_page.value:
                if not isinstance(item, DriveItem):
                    raise TypeError(
                        f"Erro ao tentar converter {item.__class__} para SharePointITEM"
                    )
                page_content.append(parse_drive_item(item))
            yield SharePointItemCollection.from_collection(page_content)

            # A pagina seguinte e solicitada apenas quando o consumidor avanca.
            current_page = await page_iterator.next()

    async def list_children(
        self,
        library: DocumentLibrary,
        parent: SharePointItem,
        filter: Callable[[SharePointItem], bool]
        | None = None,  # Deve retornar uma condição boleana
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
            # Reaproveita `list_children` para manter uma unica rota de
            # navegacao remota.
            children = await self.list_children(
                library, parent, filter=lambda item: item.name == name
            )

            # A busca e local aos filhos imediatos; nao percorre subpastas.
            for child in children:
                if child.name == name:
                    return child

            # Ausencia de filho com esse nome e resultado valido para busca.
            return None

        except NotAFolderError:
            # O pai informado nao satisfaz o contrato de navegacao.
            raise
        except DriveItemNotFoundError:
            # A falha veio da listagem base; mantemos o erro original.
            raise
        except ODataError as error:
            # O contexto da operacao ajuda o parser a diferenciar falhas de
            # pasta, item e drive.
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
            # Reaproveita `list_children` para manter uma unica rota de
            # navegacao remota.
            children = await self.list_children(
                library, parent, filter=lambda item: item.id == item_id
            )

            # A busca e local aos filhos imediatos; nao percorre subpastas.
            for child in children:
                if child.id == item_id:
                    return child

            # Nao encontrar o id entre os filhos diretos nao e erro remoto.
            return None
        except NotAFolderError:
            # O pai informado nao satisfaz o contrato de navegacao.
            raise
        except DriveItemNotFoundError:
            # A falha veio da listagem base; mantemos o erro original.
            raise
        except ODataError as error:
            # O contexto da operacao ajuda o parser a diferenciar falhas de
            # pasta, item e drive.
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
            # Reaproveita `list_children` para manter uma unica rota de
            # navegacao remota.
            children = await self.list_children(
                library, parent, filter=lambda item: item.web_url == web_url
            )

            # A busca e local aos filhos imediatos; nao percorre subpastas.
            for child in children:
                if child.web_url == web_url:
                    return child

            # URL ausente entre filhos imediatos e resultado valido de busca.
            return None
        except NotAFolderError:
            # O pai informado nao satisfaz o contrato de navegacao.
            raise
        except DriveItemNotFoundError:
            # A falha veio da listagem base; mantemos o erro original.
            raise
        except ODataError as error:
            # O contexto da operacao ajuda o parser a diferenciar falhas de
            # pasta, item e drive.
            raise parse_o_data_error(error, operation="list_children")

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
            # A busca de pasta parte da mesma listagem imediata usada pelos
            # buscadores genericos; ela nao percorre subpastas.
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
        except ODataError as e:
            parse_o_data_error(e)

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
        except ODataError as e:
            parse_o_data_error(e)

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
            # Reaproveita a listagem imediata do pai; ids de DriveItem sao
            # comparados apenas entre os filhos diretos retornados pelo Graph.
            children = await self.list_children(library, parent)

            # Um arquivo com o mesmo id quebra o contrato desta funcao, que
            # promete devolver apenas pastas.
            for child in children:
                if child.id == item_id and child.is_file:
                    raise NotAFolderError(
                        "Conflito de tipo ao buscar pasta filha pelo id "
                        f"'{item_id}': o item encontrado e um arquivo."
                    )
                # Se o id bate e nao houve conflito de arquivo, devolve o item
                # como pasta candidata.
                if child.id == item_id:
                    return child

            # Id nao encontrado entre filhos imediatos.
            return None

        # Pode vir da validacao do pai em `list_children` ou do conflito
        # de tipo detectado neste escopo.
        except NotAFolderError:
            raise

        # Falhas de item remoto ausente sao reenquadradas com o contexto de
        # navegacao usado nesta busca.
        except DriveItemNotFoundError:
            raise
        # Traduz erros remotos do SDK para erros internos do Core.
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
            # A URL web e comparada apenas entre os filhos imediatos do item
            # pai.
            children = await self.list_children(library, parent)

            # Se a URL apontar para arquivo, o resultado existe, mas nao
            # satisfaz o contrato de pasta.
            for child in children:
                if child.web_url == web_url and child.is_file:
                    raise NotAFolderError(
                        "Conflito de tipo ao buscar pasta filha pela URL "
                        f"'{web_url}': o item encontrado e um arquivo."
                    )
                # URL encontrada em um item que nao foi classificado como
                # arquivo.
                if child.web_url == web_url:
                    return child

            # URL nao encontrada entre filhos imediatos.
            return None

        # Pode vir da validacao do pai em `list_children` ou do conflito
        # de tipo detectado neste escopo.
        except NotAFolderError:
            raise

        # Falha de item remoto ausente durante a listagem de filhos.
        except DriveItemNotFoundError:
            raise
        # Traduz erros remotos do SDK para erros internos do Core.
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
            # Criacao de pasta sempre acontece dentro de outro `DriveItem` que
            # precisa representar uma pasta remota.
            if not parent.is_folder:
                raise NotAFolderError(f"{parent} não é um Folder válido")

            # O builder concentra a regra de montagem do body Graph: nome,
            # facet `folder` e comportamento de conflito.
            body = build_folder_drive_item(folder_name, conflict_behavior)

            # A rota cria um novo filho sob o item pai. O id do pai identifica
            # a posicao remota; nao ha concatenacao manual de path.
            response = (
                await self._client_manager.client.drives.by_drive_id(library.id)
                .items.by_drive_item_id(parent.id)
                .children.post(body)
            )

            # Uma criacao sem body de resposta nao consegue ser convertida para
            # `SharePointItem`, entao falha antes do parser.
            if not response:
                raise FailedWhenCreateDriveItemError(
                    "Erro ao criar o Folder especificado"
                )
        except (
            NotAFolderError,
            InvalidRemoteNameError,
            InvalidConflictBehaviorError,
            FailedWhenCreateDriveItemError,
        ):
            # Erros internos de validacao ou de envelope sao relancados sem
            # sobrescrever mensagem nem traceback.
            raise
        except ODataError as error:
            # Erros crus do SDK sao traduzidos para a hierarquia semantica do
            # Core no limite do servico.
            raise parse_o_data_error(error, operation="create_folder")

        else:
            # O SDK devolve um `DriveItem`; o Core expõe apenas a referencia
            # interna enxuta.
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
            # A navegacao sempre parte de uma pasta remota ja resolvida. Quando
            # `folders_parts` vier vazio, esta propria pasta sera retornada.
            if not root.is_folder:
                raise NotAFolderError(f"{root} não é um Folder válido")

            current_part: SharePointItem = root

            # Cada parte representa exatamente um nivel da arvore remota. A
            # funcao nao concatena ids nem caminhos; ela sempre navega a partir
            # do `SharePointItem` atual.
            for folder_part in folders_parts:
                # Primeiro tentamos reaproveitar uma pasta filha imediata
                # existente sob o item atual.
                finded_item = await self.find_folder_by_name(
                    library, current_part, folder_part
                )

                if not finded_item:
                    # Quando a pasta nao existe, criamos exatamente neste nivel
                    # e seguimos a navegacao a partir do item recem-criado.
                    created_part: SharePointItem = await self.create_folder(
                        library,
                        current_part,
                        folder_part,
                        conflict_behavior,
                    )
                    current_part = created_part
                    continue

                # Quando a pasta ja existe, ela se torna o novo ponto
                # de partida para a proxima parte do caminho.
                current_part = finded_item

            # O retorno sempre representa a pasta final da cadeia solicitada ou
            # o proprio `root` quando nenhuma parte foi informada.
            return current_part
        except (
            NotAFolderError,
            DriveItemNotFoundError,
            InvalidRemoteNameError,
            InvalidConflictBehaviorError,
            FailedWhenCreateDriveItemError,
        ):
            # A funcao e orquestradora: ela nao muda a semantica dos erros que
            # vierem das etapas de busca/criacao.
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
            raise NotAFolderError(f"{parent.name} não é um Folder válido")
        if not local_file.path.exists():
            raise LocalPathNotFoundError("O caminho para o arquivo local não existe")

        if not local_file.path.is_file():
            raise LocalPathIsDirectoryError(
                "O caminho passado aponta para um diretório"
            )

        if not local_file.size:
            raise LocalFileNotReadableError("Não foi possível ler o arquivo passado")

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
                f"O item {
                    parent.id
                } nao pode receber upload porque nao representa uma pasta."
            )

        if not local_file.path.exists():
            raise LocalPathNotFoundError(
                f"O caminho local informado nao existe: {local_file.path}"
            )
        if not local_file.path.is_file():
            raise LocalPathIsDirectoryError(
                f"Caminho local nao e um arquivo: {local_file.path}"
            )
        if local_file.size is not None and local_file.size > 250_000_000:
            raise FileVeryLargeError(
                f"O arquivo {
                    local_file.path
                } excede o limite de 250 MB do upload pequeno."
            )

        try:
            # O fluxo pequeno trabalha com o arquivo inteiro em memoria.
            # Portanto, a leitura local acontece antes do `PUT`.
            content_bytes = local_file.path.read_bytes()
        except OSError as error:
            raise LocalFileNotReadableError(
                f"Falha ao ler o arquivo local: {local_file.path}"
            ) from error

        try:
            # O staging concentra o nome remoto final e o fragmento de rota
            # usado na criacao por nome.
            staging_content: PreparedUpload = build_upload_content(
                local_file,
                None,
                conflict_behavior,
            )
            existing_file = await self.find_file_by_name(
                library,
                parent,
                local_file.name,
            )

            response = None
            remote_name = local_file.name

            if existing_file is None:
                # Sem conflito remoto, o upload pequeno cria o arquivo usando a
                # rota `items/{parent-id}:/{filename}:/content`.
                create_url = build_drive_create_content_url(
                    library.id,
                    parent.id,
                    staging_content.target_path,
                )
                response = await (
                    self._client_manager.client.drives.by_drive_id(library.id)
                    .items.by_drive_item_id(parent.id)
                    .content.with_url(create_url)
                    .put(content_bytes)
                )
            elif staging_content.conflict_behavior == "fail":
                raise FileAlreadyExistError(
                    f"Ja existe um arquivo chamado {local_file.name} na pasta remota {
                        parent.id
                    }."
                )
            elif staging_content.conflict_behavior == "rename":
                # Em modo `rename`, o Core gera um novo nome remoto e tenta a
                # criacao novamente sob o mesmo pai.
                remote_name = rename_with_uuid(local_file.name)
                staging_content = build_upload_content(
                    local_file.rename(remote_name),
                    remote_name,
                    conflict_behavior,
                )
                create_url = build_drive_create_content_url(
                    library.id,
                    parent.id,
                    staging_content.target_path,
                )
                response = await (
                    self._client_manager.client.drives.by_drive_id(library.id)
                    .items.by_drive_item_id(parent.id)
                    .content.with_url(create_url)
                    .put(content_bytes)
                )
            else:
                # Em modo `replace`, o upload usa diretamente o `item_id` do
                # arquivo existente e substitui apenas o conteudo.
                response = await (
                    self._client_manager.client.drives.by_drive_id(library.id)
                    .items.by_drive_item_id(existing_file.id)
                    .content.put(content_bytes)
                )
                remote_name = existing_file.name or local_file.name

            if not response:
                raise SmallFileUploadError(
                    "O Graph nao retornou um DriveItem apos o upload pequeno "
                    f"de {local_file.path}."
                )

            # O retorno do SDK ainda e um `DriveItem`; o parser o reduz para a
            # referencia semantica usada pelo restante do Core.
            uploaded_item = parse_drive_item(response)
            return FileUploadResult(
                item=uploaded_item,
                source_path=local_file.path,
                conflict_behavior=staging_content.conflict_behavior,
                remote_name=remote_name,
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
            raise parse_o_data_error(error, operation="upload_small_file")

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
        # O endpoint de sessao cria o arquivo como filho de um item pasta.
        if not parent.is_folder:
            raise NotAFolderError(f"O item {parent.id} nao e uma pasta de destino.")

        # As validacoes locais evitam criar uma sessao que nao podera ser
        # usada.
        if not local_file.path.exists():
            raise LocalPathNotFoundError(
                f"Caminho local invalido para arquivo: {local_file.path}"
            )

        if not local_file.path.is_file():
            raise LocalPathIsDirectoryError(
                f"O caminho {local_file.path} aponta para um diretório"
            )

        try:
            upload_session = await self._create_upload_session(
                library=library,
                parent=parent,
                local_file=local_file,
                conflict_behavior=conflict_behavior,
            )

            # O context manager garante o fechamento do descritor local ao fim
            # do envio ou quando uma excecao interromper o fluxo.
            with local_file.path.open("rb") as file_stream:
                # A task usa a sessao pre-autenticada para controlar os
                # ranges e enviar cada trecho do stream.
                upload_task = LargeFileUploadTask(
                    upload_session=upload_session,
                    request_adapter=(self._client_manager.client.request_adapter),
                    stream=file_stream,  # type: ignore
                    parsable_factory=DriveItem,
                    max_chunk_size=MAX_CHUNK_SIZE,
                )
                upload_result = await upload_task.upload()

            # O resultado cru do SDK informa se a task considera o envio
            # concluido.
            if not upload_result.upload_succeeded:
                raise RuntimeError(f"Upload nao concluido para {local_file.name}")
        except LocalPathNotFoundError:
            raise
        except LocalPathIsDirectoryError:
            raise
        except RuntimeError:
            raise
        if not upload_result.item_response:
            raise UploadError("Não foi possível completar o Upload")

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
        # O helper repete as validacoes para continuar seguro caso passe a ser
        # chamado por outro fluxo interno.
        if not local_file.path.exists():
            raise LocalPathNotFoundError(
                f"Caminho local invalido para arquivo: {local_file.path}"
            )

        if not local_file.path.is_file():
            raise LocalPathIsDirectoryError(
                f"O caminho {local_file.path} aponta para um diretório"
            )

        if local_file.size == 0:
            raise LocalFileNotReadableError(
                f"O caminho {local_file.path} aponta para um caminho vazio"
            )

        # As propriedades uploadable carregam metadados aplicados ao arquivo
        # remoto quando a sessao for finalizada.
        uploadable_propieties = DriveItemUploadableProperties(
            additional_data={"@microsoft.graph.conflictBehavior": conflict_behavior}
        )

        request_body = CreateUploadSessionPostRequestBody(item=uploadable_propieties)

        # O endereco por path combina o id do pai com o nome do arquivo
        # que sera criado ou resolvido pelo Graph.
        upload_session = await (
            self._client_manager.client.drives.by_drive_id(library.id)
            .items.by_drive_item_id(f"{parent.id}:/{local_file.name}:")
            .create_upload_session.post(request_body)
        )

        # A task de upload precisa obrigatoriamente da URL pre-autenticada.
        if upload_session is None:
            raise RuntimeError("O Microsoft Graph não retornou uma URL válida")

        if upload_session.upload_url is None:
            raise RuntimeError("A sessão retornada não possui url")

        return upload_session
