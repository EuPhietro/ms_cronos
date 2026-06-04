'''Servico de alto nivel para operacoes do SharePoint via Microsoft Graph.

Este modulo concentra a adaptacao entre o SDK do Graph e os contratos semanticos
do Core. O objetivo e expor operacoes pequenas, previsiveis e tipadas, sem
vazar envelopes crus do SDK para as camadas superiores.

Exemplo de uso:
    service = SharePointService(graph_client_manager)
    site = await service.resolve_site('https://tenant.sharepoint.com/sites/RH')
    drive = await service.get_default_drive(site)
    root = await service.get_drive_root(drive)
    children = await service.list_children(drive, root)
'''

from __future__ import annotations

from typing import Literal, Optional, Sequence

from msgraph.generated.models.o_data_errors.o_data_error import ODataError

from core import (
    adapt_site,
    build_folder_drive_item,
    build_drive_create_content_url,
    build_graph_site_url,
    DefaultDriveNotFoundError,
    DriveItemCollection,
    DriveItemNotFoundError,
    DriveItemRef,
    parse_drive,
    parse_drive_collection_response,
    parse_drive_item,
    parse_drive_item_collection_response,
    parse_o_data_error,
    DriveRef,
    DriveRefCollection,
    FailedWhenCreateDriveItemError,
    GraphClientManager,
    GraphResponseError,
    DriveNotFoundError,
    InvalidConflictBehaviorError,
    InvalidRemoteNameError,
    NotAFolderError,
    parse_site,
    SiteRef,
    SiteResolutionError,
)
from core.builders import build_upload_content
from core.errors import (
    FileAlreadyExistError,
    FileVeryLargeError,
    LocalFileNotReadableError,
    LocalPathIsDirectoryError,
    LocalPathNotFoundError,
    NotAFileError,
    SmallFileUploadError,
)
from core.models import LocalFile, StagingContentUpload, UploadResult
from core.utils import rename_with_uuid


class SharePointService:
    '''Servico principal do Core para leitura de sites, drives e itens.

    A classe orquestra chamadas ao SDK, valida envelopes minimos e delega a
    conversao dos models crus ao modulo `parse`.
    '''

    def __init__(self, graph_client_manager: GraphClientManager) -> None:
        # O manager encapsula autenticacao e ciclo de vida do client Graph,
        # deixando o servico focado apenas em regras de SharePoint.
        self._client_manager = graph_client_manager

    async def resolve_site(
        self,
        sharepoint_url: str,
    ) -> SiteRef:
        '''Resolve uma URL humana do SharePoint e devolve um `SiteRef`.

        A rota `sites.with_url(...)` pode retornar dados em formatos diferentes
        dependendo do SDK; por isso este metodo aceita tanto `response.value`
        quanto `response.additional_data`.
        '''
        try:
            # A URL humana e convertida para a rota Graph que entende
            # `hostname:/server-relative-path`.
            secure_url = build_graph_site_url(sharepoint_url)

            # A rota `sites.with_url(...)` resolve o site sem exigir que o
            # chamador conheca o `site_id` antes.
            response = await self._client_manager.client.sites.with_url(secure_url).get()
            
            # Sem envelope, o Core nao consegue distinguir site inexistente de
            # resposta remota inconsistente.
            if not response:
                raise SiteResolutionError(
                    f'Nao foi possivel obter resposta do Graph ao resolver o site: {sharepoint_url}'
                )

            # 4. A resolucao de site e um caso especial do SDK: em alguns
            #    cenarios o site resolvido vem em `response.value[0]`; em outros,
            #    os dados minimos chegam em `response.additional_data`.
            if response.value:
                # Quando o SDK preenche `value`, a adaptacao segue o parser
                # normal de `Site`.
                data = parse_site(response.value[0])
            elif response.additional_data.get('id'):
                # Alguns cenarios populam apenas `additional_data`; neste caso o
                # Core faz um fallback controlado para nao perder a resolucao.
                data = adapt_site(response.additional_data)
            else:
                raise GraphResponseError(
                    f'O Graph respondeu a resolucao do site, mas nao retornou dados suficientes para montar um SiteRef: {sharepoint_url}'
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
            raise parse_o_data_error(error, operation='resolve_site')
        else:
            # O retorno final expoe apenas a referencia semantica do Core.
            return data



    async def list_site_drives(
        self,
        site_ref: SiteRef,
    ) -> DriveRefCollection:
        '''Lista as bibliotecas de documentos associadas a um site resolvido.'''
        try:
            # 1. Parte de um site ja resolvido e usa apenas o `site_ref.id` para
            #    consultar as bibliotecas/document libraries associadas.
            response = await self._client_manager.client.sites.by_site_id(site_ref.id).drives.get()

            # 2. Garante que o SDK devolveu o envelope da operacao.
            if not response:
                raise DriveNotFoundError(
                    f'O Graph nao retornou resposta ao listar os drives do site {site_ref.id}.'
                )

            # 3. O Core espera a lista concreta de drives dentro de
            #    `response.value`.
            if not response.value:
                raise GraphResponseError(
                    f'O Graph respondeu a listagem de drives do site {site_ref.id}, mas o envelope veio sem itens em `value`.'
                )

        except DriveNotFoundError:
            # Falhas locais de envelope vazio sao preservadas como erro interno.
            raise
        except ODataError as error:
            # O parser recebe o contexto da operacao para mapear melhor erros
            # como `itemNotFound` e `resourceNotFound`.
            raise parse_o_data_error(error, operation='list_site_drives')
        else:
            # 4. O envelope cru do SDK e adaptado para a colecao semantica do
            #    Core antes de sair do servico.
            return parse_drive_collection_response(response)

    async def get_default_drive(
        self,
        site_ref: SiteRef,
    ) -> DriveRef:
        '''Obtem o drive padrao do site a partir de um `SiteRef` ja resolvido.'''

        try:
            # 1. Usa o `site_ref.id` para consultar diretamente o drive padrao do
            #    site, sem listar todas as bibliotecas antes.
            response = await self._client_manager.client.sites.by_site_id(site_ref.id).drive.get()

            # 2. Nesta rota o Graph devolve um `Drive` unico, nao um envelope de
            #    colecao com `value`.
            if not response:
                raise DefaultDriveNotFoundError(
                    f'O Graph nao retornou o drive padrao do site {site_ref.id}.'
                )

        except ODataError as error:
            # O parser centraliza a traducao para o erro semantico correto do
            # Core.
            raise parse_o_data_error(error, operation='get_default_drive')
        else:
            # 3. O `Drive` cru retornado pelo SDK e convertido para o contrato
            #    interno enxuto antes de sair do servico.
            return parse_drive(response)

    async def get_drive(
        self,
        drive_ref: DriveRef,
    ) -> DriveRef:
        '''Reconsulta um drive especifico a partir de um `DriveRef`.

        Use quando voce ja tem um id de drive e quer confirmar ou atualizar os
        metadados basicos desse drive.
        '''
        try:
            # 1. Reconsulta um drive especifico a partir do `drive_ref.id`.
            response = await self._client_manager.client.drives.by_drive_id(drive_ref.id).get()

            # 2. Assim como em `get_default_drive`, aqui a resposta esperada e
            #    um `Drive` unico.
            if not response:
                raise DriveNotFoundError(
                    f'O Graph nao retornou o drive solicitado para o identificador {drive_ref.id}.'
                )

        except DriveNotFoundError:
            # Se a ausencia foi detectada pelo proprio Core, relancamos a mesma
            # excecao para nao apagar a origem do erro.
            raise
        except ODataError as error:
            # O parser centraliza a traducao para o erro semantico correto do
            # Core.
            raise parse_o_data_error(error, operation='get_drive')
        else:
            # 3. Converte o `Drive` cru do SDK para `DriveRef`.
            return parse_drive(response)

    async def get_drive_root(
        self,
        drive_ref: DriveRef,
    ) -> DriveItemRef:
        '''Retorna o `DriveItemRef` que representa a raiz do drive.'''
        try:
            # 1. Consulta o item raiz do drive, que no Graph e modelado como um
            #    `DriveItem`.
            site_response = await self._client_manager.client.drives.by_drive_id(drive_ref.id).root.get()

            # 2. Nesta rota a resposta ja e um item unico, nao um envelope de
            #    colecao.
            if not site_response:
                raise DriveNotFoundError(
                    f'O Graph nao retornou a raiz do drive {drive_ref.id}.'
                )

        except DriveNotFoundError:
            # A raiz do drive e obrigatoria para navegacao; se nao veio resposta,
            # o erro interno ja descreve o drive afetado.
            raise
        except ODataError as error:
            # O parser centraliza a traducao para o erro semantico correto do
            # Core.
            raise parse_o_data_error(error, operation='get_drive_root')
        else:
            # 3. Converte o `DriveItem` raiz em `DriveItemRef`.
            return parse_drive_item(site_response)

    async def list_children(
        self,
        drive_ref: DriveRef,
        parent_item_ref: DriveItemRef,
    ) -> DriveItemCollection:
        '''Lista os filhos imediatos de uma pasta remota.

        A busca nao e recursiva: cada chamada representa exatamente um nivel da
        arvore remota.
        '''

        try:
            # 1. Parte de um drive e de um item pai ja resolvidos para consultar
            #    os filhos imediatos da pasta no Graph.
            if not parent_item_ref.is_folder:
                raise NotAFolderError(
                    f'O item {parent_item_ref.id} nao pode ser usado em list_children porque nao representa uma pasta.'
                )

            response = await self._client_manager.client.drives.by_drive_id(drive_ref.id).items.by_drive_item_id(parent_item_ref.id).children.get()

            # 2. A rota de filhos devolve um envelope de colecao contendo
            #    `DriveItem`s em `response.value`.
            if not response:
                raise DriveItemNotFoundError(
                    f'O Graph nao retornou resposta ao listar os filhos do item {parent_item_ref.id} no drive {drive_ref.id}.'
                )

        except NotAFolderError:
            # Um item que nao e pasta nao pode ter filhos navegaveis neste
            # contrato.
            raise
        except DriveItemNotFoundError:
            # Envelope ausente ou item inacessivel continua como erro interno de
            # DriveItem.
            raise
        except ODataError as error:
            # O contexto da operacao ajuda o parser a diferenciar falhas de
            # pasta, item e drive.
            raise parse_o_data_error(error, operation='list_children')
        else:
            # 3. O envelope cru e convertido para `DriveItemCollection`, que e a
            #    forma semantica de navegacao usada pelo Core.
            return parse_drive_item_collection_response(response)

    async def find_child_by_name(
        self,
        drive_ref: DriveRef,
        parent_item_ref: DriveItemRef,
        name: str,
    ) -> Optional[DriveItemRef]:
        '''Busca um filho imediato pelo nome dentro de uma pasta remota.'''
        try:
            # Reaproveita `list_children` para manter uma unica rota de
            # navegacao remota.
            children = await self.list_children(drive_ref, parent_item_ref)

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
            raise parse_o_data_error(error, operation='list_children')

    async def find_child_by_id(
        self,
        drive_ref: DriveRef,
        parent_item_ref: DriveItemRef,
        drive_id: str,
    ) -> Optional[DriveItemRef]:
        '''Busca um filho imediato pelo id dentro de uma pasta remota.'''
        try:
            # Reaproveita `list_children` para manter uma unica rota de
            # navegacao remota.
            children = await self.list_children(drive_ref, parent_item_ref)

            # A busca e local aos filhos imediatos; nao percorre subpastas.
            for child in children:
                if child.id == drive_id:
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
            raise parse_o_data_error(error, operation='list_children')

    async def find_child_by_web_url(
        self,
        drive_ref: DriveRef,
        parent_item_ref: DriveItemRef,
        web_url: str,
    ) -> Optional[DriveItemRef]:
        '''Busca um filho imediato pela URL web dentro de uma pasta remota.'''
        try:
            # Reaproveita `list_children` para manter uma unica rota de
            # navegacao remota.
            children = await self.list_children(drive_ref, parent_item_ref)

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
            raise parse_o_data_error(error, operation='list_children')

    async def find_child_folder(
        self,
        drive_ref: DriveRef,
        parent_item_ref: DriveItemRef,
        name: str,
    ) -> Optional[DriveItemRef]:
        '''Busca uma pasta filha imediata pelo nome.'''
        try:
            # A busca de pasta parte da mesma listagem imediata usada pelos
            # buscadores genericos; ela nao percorre subpastas.
            child = await self.find_child_by_name(drive_ref, parent_item_ref, name=name)

            # Cada filho encontrado pode ser arquivo ou pasta. Se o nome bate com
            # um arquivo, isso e um conflito de tipo para este contrato.
            if not child:
                return None

            if child.is_file:
                raise NotAFolderError(
                    f"Conflito de tipo ao buscar pasta filha '{name}': existe um arquivo com esse nome no item pai {parent_item_ref.id}."
                )
            # Se o nome bate e nao houve conflito de arquivo, o item pode ser
            # usado como pasta remota de destino/navegacao.

            return child
        # `NotAFolderError` pode vir do pai informado a `list_children` ou do
        # conflito local em que existe arquivo com o mesmo nome esperado.
        except NotAFolderError:
            raise

        # Erros de item/drive ausente preservam contexto de pai e drive para
        # facilitar depuracao da navegacao remota.
        except DriveItemNotFoundError:
            raise
        # Erros OData crus do SDK sao traduzidos para a hierarquia semantica do
        # Core antes de atravessar a fronteira do servico.
        except ODataError as error:
            raise parse_o_data_error(error, operation='find_child_folder')

    async def find_child_file(
        self,
        drive_ref: DriveRef,
        parent_item_ref: DriveItemRef,
        name: str,
    ) -> Optional[DriveItemRef]:
        '''Busca um arquivo filho imediato pelo nome.'''
        try:
            # A busca de arquivo reaproveita a mesma listagem imediata usada
            # pelos demais buscadores; ela nao percorre subpastas.
            child = await self.find_child_by_name(drive_ref, parent_item_ref, name=name)

            # Cada filho encontrado pode ser arquivo ou pasta. Se o nome bate com
            # uma pasta, isso e um conflito de tipo para este contrato.
            if not child:
                return None

            if child.is_folder:
                raise NotAFileError(
                    f"Conflito de tipo ao buscar arquivo filho '{name}': existe uma pasta com esse nome no item pai {parent_item_ref.id}."
                )
            # Se o nome bate e nao houve conflito de pasta, o item pode ser
            # usado como arquivo remoto de destino/substituicao.

            return child
        # `NotAFolderError` pode vir do pai informado a `list_children` ou do
        # conflito local em que existe arquivo com o mesmo nome esperado.
        except NotAFolderError:
            raise

        # Erros de item/drive ausente preservam contexto de pai e drive para
        # facilitar depuracao da navegacao remota.
        except DriveItemNotFoundError:
            raise
        # Erros OData crus do SDK sao traduzidos para a hierarquia semantica do
        # Core antes de atravessar a fronteira do servico.
        except ODataError as error:
            raise parse_o_data_error(error, operation='find_child_file')

    async def find_child_folder_by_id(
        self,
        drive_ref: DriveRef,
        parent_item_ref: DriveItemRef,
        drive_id: str,
    ) -> Optional[DriveItemRef]:
        '''Busca uma pasta filha imediata pelo id do DriveItem.'''
        try:
            # Reaproveita a listagem imediata do pai; ids de DriveItem sao
            # comparados apenas entre os filhos diretos retornados pelo Graph.
            children = await self.list_children(drive_ref, parent_item_ref)

            # Um arquivo com o mesmo id procurado quebra o contrato desta funcao,
            # que promete devolver apenas pastas.
            for child in children:

                if child.id == drive_id and child.is_file:

                    raise NotAFolderError(
                        f"Conflito de tipo ao buscar pasta filha pelo id '{drive_id}': \
                        o item encontrado é um arquivo, não uma pasta."
                    )
                # Se o id bate e nao houve conflito de arquivo, devolve o item
                # como pasta candidata.
                if child.id == drive_id:
                    return child

            # Id nao encontrado entre filhos imediatos.
            return None

        # Pode vir da validacao do pai em `list_children` ou do conflito de tipo
        # detectado neste escopo.
        except NotAFolderError:
            raise

        # Falhas de item remoto ausente sao reenquadradas com o contexto de
        # navegacao usado nesta busca.
        except DriveItemNotFoundError:
            raise
        # Traduz erros remotos do SDK para erros internos do Core.
        except ODataError as error:
            raise parse_o_data_error(error, operation='find_child_folder')

    async def find_child_folder_web_url(
        self,
        drive_ref: DriveRef,
        parent_item_ref: DriveItemRef,
        web_url: str,
    ) -> Optional[DriveItemRef]:
        '''Busca uma pasta filha imediata pela URL web do DriveItem.'''
        try:
            # A URL web e comparada apenas entre os filhos imediatos do item pai.
            children = await self.list_children(drive_ref, parent_item_ref)

            # Se a URL apontar para arquivo, o resultado existe mas nao satisfaz
            # o contrato de pasta.
            for child in children:

                if child.web_url == web_url and child.is_file:

                    raise NotAFolderError(
                        f"Conflito de tipo ao buscar pasta filha pela URL '{web_url}': o item encontrado é um arquivo, não uma pasta."
                    )
                # URL encontrada em um item que nao foi classificado como arquivo.
                if child.web_url == web_url:
                    return child

            # URL nao encontrada entre filhos imediatos.
            return None

        # Pode vir da validacao do pai em `list_children` ou do conflito de tipo
        # detectado neste escopo.
        except NotAFolderError:
            raise

        # Falha de item remoto ausente durante a listagem de filhos.
        except DriveItemNotFoundError:
            raise
        # Traduz erros remotos do SDK para erros internos do Core.
        except ODataError as error:
            raise parse_o_data_error(error, operation='find_child_folder')

    async def create_folder(
        self,
        drive_ref: DriveRef,
        parent_item_ref: DriveItemRef,
        folder_name: str,
        conflict_behavior: Literal['fail', 'rename', 'replace'] = 'fail'
    ) -> DriveItemRef:
        '''Cria uma pasta filha imediata dentro de uma pasta remota.'''

        try:
            # Criacao de pasta sempre acontece dentro de outro `DriveItem` que
            # precisa representar uma pasta remota.
            if not parent_item_ref.is_folder:
                raise NotAFolderError(
                    f'{parent_item_ref} não é um Folder válido')

            # O builder concentra a regra de montagem do body Graph: nome,
            # facet `folder` e comportamento de conflito.
            body = build_folder_drive_item(folder_name, conflict_behavior)

            # A rota cria um novo filho sob o item pai informado. O id do pai ja
            # identifica a posicao remota; nao ha concatenacao manual de path.
            response = await self._client_manager.client.drives.by_drive_id(drive_ref.id) \
                .items.by_drive_item_id(parent_item_ref.id).children.post(body)

            # Uma criacao sem body de resposta nao consegue ser convertida para
            # `DriveItemRef`, entao falha antes do parser.
            if not response:
                raise FailedWhenCreateDriveItemError(
                    'Erro ao criar o Folder especificado')
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
            raise parse_o_data_error(error, operation='create_folder')

        else:
            # O SDK devolve um `DriveItem`; o Core expõe apenas a referencia
            # interna enxuta.
            return parse_drive_item(response)

    async def ensure_remote_folder_path(
        self,
        drive_ref: DriveRef,
        root_item: DriveItemRef,
        folders_parts: Sequence[str],
        conflict_behavior: Literal['fail', 'rename', 'replace'] = 'fail'
    ) -> DriveItemRef:
        '''Garante uma cadeia de pastas remotas a partir de uma pasta raiz.'''

        try:
            # A navegacao sempre parte de uma pasta remota ja resolvida. Quando
            # `folders_parts` vier vazio, esta propria pasta sera retornada.
            if not root_item.is_folder:
                raise NotAFolderError(f'{root_item} não é um Folder válido')

            current_part: DriveItemRef = root_item

            # Cada parte representa exatamente um nivel da arvore remota. A
            # funcao nao concatena ids nem caminhos; ela sempre navega a partir
            # do `DriveItemRef` atual.
            for folder_part in folders_parts:
                # Primeiro tentamos reaproveitar uma pasta filha imediata que ja
                # exista sob o item atual.
                finded_item = await self.find_child_folder(drive_ref, current_part, folder_part)

                if not finded_item:
                    # Quando a pasta nao existe, criamos exatamente neste nivel
                    # e seguimos a navegacao a partir do item recem-criado.
                    created_part: DriveItemRef = await self.create_folder(
                        drive_ref,
                        current_part,
                        folder_part,
                        conflict_behavior,
                    )
                    current_part = created_part
                    continue

                # Quando a pasta ja existe, ela se torna o novo ponto de partida
                # para a proxima parte do caminho.
                current_part = finded_item

            # O retorno sempre representa a pasta final da cadeia solicitada ou
            # o proprio `root_item` quando nenhuma parte foi informada.
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

    async def upload_small_file(
        self,
        drive_ref: DriveRef,
        parent_item_ref: DriveItemRef,
        local_file: LocalFile,
        conflict_behavior: Literal["fail", "rename", "replace"] = 'fail',
        ) -> UploadResult:
        """Envia um arquivo pequeno por PUT direto no endpoint `/content`.

        O fluxo decide entre criar ou substituir de acordo com a existencia do
        arquivo remoto e com o `conflict_behavior` informado.
        """
        if not parent_item_ref.is_folder:
            raise NotAFolderError(
                f'O item {parent_item_ref.id} nao pode receber upload porque nao representa uma pasta.'
            )

        if not local_file.path.exists():
            raise LocalPathNotFoundError(
                f'O caminho local informado nao existe: {local_file.path}'
            )
        if not local_file.path.is_file():
            raise LocalPathIsDirectoryError(
                f'O caminho local precisa apontar para um arquivo, mas recebeu: {local_file.path}'
            )
        if local_file.size is not None and local_file.size > 250_000_000:
            raise FileVeryLargeError(
                f'O arquivo {local_file.path} excede o limite de 250 MB do upload pequeno.'
            )

        try:
            # O fluxo pequeno trabalha com o arquivo inteiro em memoria, entao a
            # leitura local precisa acontecer antes do `PUT`.
            content_bytes = local_file.path.read_bytes()
        except OSError as error:
            raise LocalFileNotReadableError(
                f'Nao foi possivel ler o arquivo local para upload: {local_file.path}'
            ) from error

        try:
            # O staging concentra o nome remoto final e o fragmento de rota
            # usado na criacao por nome.
            staging_content: StagingContentUpload = build_upload_content(
                local_file,
                None,
                conflict_behavior,
            )
            existing_file = await self.find_child_file(
                drive_ref,
                parent_item_ref,
                local_file.name,
            )

            response = None
            remote_name = local_file.name

            if existing_file is None:
                # Sem conflito remoto, o upload pequeno cria o arquivo usando a
                # rota `items/{parent-id}:/{filename}:/content`.
                create_url = build_drive_create_content_url(
                    drive_ref.id,
                    parent_item_ref.id,
                    staging_content.target_path,
                )
                response = await (
                    self._client_manager.client.drives
                    .by_drive_id(drive_ref.id)
                    .items.by_drive_item_id(parent_item_ref.id)
                    .content.with_url(create_url)
                    .put(content_bytes)
                )
            elif staging_content.conflict_behavior == 'fail':
                raise FileAlreadyExistError(
                    f'Ja existe um arquivo chamado {local_file.name} na pasta remota {parent_item_ref.id}.'
                )
            elif staging_content.conflict_behavior == 'rename':
                # Em modo `rename`, o Core gera um novo nome remoto e tenta a
                # criacao novamente sob o mesmo pai.
                remote_name = rename_with_uuid(local_file.name)
                staging_content = build_upload_content(
                    local_file.rename(remote_name),
                    remote_name,
                    conflict_behavior,
                )
                create_url = build_drive_create_content_url(
                    drive_ref.id,
                    parent_item_ref.id,
                    staging_content.target_path,
                )
                response = await (
                    self._client_manager.client.drives
                    .by_drive_id(drive_ref.id)
                    .items.by_drive_item_id(parent_item_ref.id)
                    .content.with_url(create_url)
                    .put(content_bytes)
                )
            else:
                # Em modo `replace`, o upload usa diretamente o `item_id` do
                # arquivo existente e substitui apenas o conteudo.
                response = await (
                    self._client_manager.client.drives
                    .by_drive_id(drive_ref.id)
                    .items.by_drive_item_id(existing_file.id)
                    .content.put(content_bytes)
                )
                remote_name = existing_file.name or local_file.name

            if not response:
                raise SmallFileUploadError(
                    f'O Graph nao retornou um DriveItem apos o upload pequeno de {local_file.path}.'
                )

            # O retorno do SDK ainda e um `DriveItem`; o parser o reduz para a
            # referencia semantica usada pelo restante do Core.
            uploaded_item = parse_drive_item(response)
            return UploadResult(
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
            raise parse_o_data_error(error, operation='upload_small_file')
