from __future__ import annotations

'''Servico de alto nivel para operacoes do SharePoint via Microsoft Graph.

Este modulo concentra a adaptacao entre o SDK do Graph e os contratos semanticos
do Core. O objetivo e expor operacoes pequenas, previsiveis e tipadas, sem
vazar envelopes crus do SDK para as camadas superiores.
'''

from msgraph.generated.models.o_data_errors.o_data_error import ODataError

from core import (
    GraphClientManager,
    SiteRef,
    SiteRefCollection,
    DriveRef,
    DriveRefCollection,
    DriveItemRef,
    DriveItemCollection,
    LocalFile,
    LocalFileCollection,
    build_graph_site_url,
    parse_site,
    adapt_site,
    parse_drive,
    parse_drive_item,
    parse_local_file,
    parse_site_collection_response,
    parse_drive_collection_response,
    parse_drive_item_collection_response,
    parse_o_data_error,
    SiteResolutionError,
    GraphResponseError,
    DefaultDriveNotFoundError,
    DriveNotFoundError,
    NotAFolderError,
    NotAFileError,
    DriveItemNotFoundError
)


class SharePointService:
    '''Servico principal do Core para leitura de sites, drives e itens.'''
    
    def __init__(self, graph_client_manager: GraphClientManager) -> None:
        self._client_manager = graph_client_manager

    async def resolve_site(self, sharepoint_url: str) -> SiteRef:
        '''Resolve uma URL humana do SharePoint e devolve um `SiteRef`.'''
        try:
            # 1. Converte a URL humana do SharePoint na rota de resolucao do
            #    Graph que aceita hostname + path.
            secure_url = build_graph_site_url(sharepoint_url)

            # 2. Consulta o Graph para descobrir o site correspondente.
            response = await self._client_manager.client.sites.with_url(secure_url).get()

            # 3. Falha cedo se o SDK nao devolver nenhum envelope de resposta.
            if not response:
                raise SiteResolutionError(
                    f'Nao foi possivel obter resposta do Graph ao resolver o site: {sharepoint_url}'
                )

            # 4. A resolucao de site e um caso especial do SDK: em alguns
            #    cenarios o site resolvido vem em `response.value[0]`; em outros,
            #    os dados minimos chegam em `response.additional_data`.
            if response.value:
                # 4.1. Quando o envelope vier preenchido com `value`, o fluxo
                #      segue pelo parser padrao de `Site`.
                data = parse_site(response.value[0])
            elif response.additional_data.get('id'):
                # 4.2. Fallback especifico para a rota `sites.with_url(...)`,
                #      onde o SDK pode popular apenas `additional_data`.
                data = adapt_site(response.additional_data)
            else: 
                raise GraphResponseError(
                    f'O Graph respondeu a resolucao do site, mas nao retornou dados suficientes para montar um SiteRef: {sharepoint_url}'
                    )

        except SiteResolutionError:
            raise SiteResolutionError(
                f'Falha ao resolver o site do SharePoint a partir da URL informada: {sharepoint_url}'
            )
        except GraphResponseError:
            raise GraphResponseError(
                f'A resposta do Graph para a URL {sharepoint_url} nao continha dados suficientes para resolver o site.'
            )
        except ODataError as error:
           # O parser centraliza a traducao do erro remoto para a hierarquia
           # interna do Core.
           raise parse_o_data_error(error, operation='resolve_site')
        else:
            # 5. So depois das validacoes o model cru do SDK e convertido para o
            #    contrato interno enxuto do Core.
            return data
  
    async def list_site_drives(self,site_ref:SiteRef) -> DriveRefCollection:
        
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

        except DriveNotFoundError   as error:
            raise DriveNotFoundError(
                f'Nao foi possivel listar os drives do site {site_ref.id}.'
            )
        except ODataError as error:
           # O parser recebe o contexto da operacao para mapear melhor erros
           # como `itemNotFound` e `resourceNotFound`.
           raise parse_o_data_error(error, operation='list_site_drives')
        else:
            # 4. O envelope cru do SDK e adaptado para a colecao semantica do
            #    Core antes de sair do servico.
            return parse_drive_collection_response(response)
           
    async def get_default_drive(self,site_ref:SiteRef) -> DriveRef:
        
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
    
    async def get_drive(self,drive_ref:DriveRef) -> DriveRef:
        
        try:
            # 1. Reconsulta um drive especifico a partir do `drive_ref.id`.
            response = await self._client_manager.client.drives.by_drive_id(drive_ref.id).get()

            # 2. Assim como em `get_default_drive`, aqui a resposta esperada e
            #    um `Drive` unico.
            if not response:
                raise DriveNotFoundError(
                    f'O Graph nao retornou o drive solicitado para o identificador {drive_ref.id}.'
                )

        except DriveNotFoundError   as error:
            raise DriveNotFoundError(
                f'Falha ao resolver o drive a partir do identificador informado: {drive_ref.id}'
            )
        except ODataError as error:
           # O parser centraliza a traducao para o erro semantico correto do
           # Core.
           raise parse_o_data_error(error, operation='get_drive')
        else:
            # 3. Converte o `Drive` cru do SDK para `DriveRef`.
            return parse_drive(response)
        
    async def get_drive_root(self,drive_ref:DriveRef) -> DriveItemRef:
        
        try:
            # 1. Consulta o item raiz do drive, que no Graph e modelado como um
            #    `DriveItem`.
            response = await self._client_manager.client.drives.by_drive_id(drive_ref.id).root.get()

            # 2. Nesta rota a resposta ja e um item unico, nao um envelope de
            #    colecao.
            if not response:
                raise DriveNotFoundError(
                    f'O Graph nao retornou a raiz do drive {drive_ref.id}.'
                )

        except DriveNotFoundError   as error:
            raise DriveNotFoundError(
                f'Falha ao resolver a raiz do drive {drive_ref.id}.'
            )
        except ODataError as error:
           # O parser centraliza a traducao para o erro semantico correto do
           # Core.
           raise parse_o_data_error(error, operation='get_drive_root')
        else:
            # 3. Converte o `DriveItem` raiz em `DriveItemRef`.
            return parse_drive_item(response)

    async def list_children(self,drive_ref:DriveRef, parent_item_ref: DriveItemRef) -> DriveItemCollection:
        
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


        except NotAFolderError   as error:
            raise NotAFolderError(
                f'O item {parent_item_ref.id} nao e uma pasta e, por isso, nao possui filhos para listagem.'
            )
        except DriveItemNotFoundError as error:
            raise DriveItemNotFoundError(
                f'Nao foi possivel localizar os filhos do item {parent_item_ref.id} no drive {drive_ref.id}.'
            )
        except ODataError as error:
           # O contexto da operacao ajuda o parser a diferenciar falhas de
           # pasta, item e drive.
           raise parse_o_data_error(error, operation='list_children')
        else:
            # 3. O envelope cru e convertido para `DriveItemCollection`, que e a
            #    forma semantica de navegacao usada pelo Core.
            return parse_drive_item_collection_response(response)
