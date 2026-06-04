"""Laboratorio manual para validar o Core contra o Microsoft Graph.

Este arquivo nao faz parte da API do Core. Ele serve como roteiro executavel para
testar manualmente credenciais, resolucao de site, navegacao por drive e leitura
da raiz enquanto a implementacao evolui.
"""

import asyncio
from os import getenv

from dotenv import load_dotenv
from rich import print
from rich.traceback import install

from core import GraphClientManager, GraphCredentials, LocalFile, SharePointService


load_dotenv()
install(show_locals=True)


async def main() -> None:
    # Credenciais da aplicacao carregadas do ambiente local para autenticar no
    # Microsoft Graph durante os testes manuais.
    client_id = getenv('CLIENT_ID', '')
    client_secret = getenv('CLIENT_SECRET', '')
    tenant_id = getenv('CLIENT_TENANT', '')

    # Instanciação do Core
    credentials = GraphCredentials(client_id, client_secret, tenant_id)
    site_url = 'https://plangeconcombr.sharepoint.com/sites/RHConecta'
    graph_client_manager = GraphClientManager(credentials)
    sharepoint = SharePointService(graph_client_manager)
    
    try:
        # 1. Resolve a URL humana do SharePoint e obtem o `SiteRef` que sera
        # usado nas chamadas seguintes.
        site = await sharepoint.resolve_site(site_url)
        print(site)

        # 2. Lista as bibliotecas/drives disponiveis no site resolvido.
        drives_library = await sharepoint.list_site_drives(site)
        print(drives_library)

        # 3. Recupera a document library padrao do site.
        default_drive = await sharepoint.get_default_drive(site)
        print(default_drive)

        # 4. A raiz do drive funciona como ponto de partida da navegacao remota.
        root_item = await sharepoint.get_drive_root(default_drive)
        print(root_item)

        # 5. Lista os filhos imediatos da raiz para inspecionar a estrutura do
        # drive.
        children = await sharepoint.list_children(default_drive, root_item)
        print(children)

        # 6. Escolhe uma pasta filha para continuar o laboratorio.
        curriculos_drive_item = children[0]
        print(curriculos_drive_item)

        # 7. Navega para o proximo nivel da arvore remota.
        curriculos_items = await sharepoint.list_children(default_drive, curriculos_drive_item)
        print(curriculos_items)

        # 8. Cria uma pasta filha imediata sob a pasta escolhida.
        created_drive_item = await sharepoint.create_folder(
            default_drive,
            curriculos_drive_item,
            'Pasta_de_Teste_1345',
            conflict_behavior='rename',
        )
        print(created_drive_item)

        # 9. Garante uma cadeia inteira de diretorios remotos a partir da raiz.
        created_hierarcly = await sharepoint.ensure_remote_folder_path(
            default_drive,
            root_item,
            folders_parts=('datasets3', '2026', '05', '21', '20-00', '2', '0'),
        )
        print(created_hierarcly)

        # 10. Converte um caminho local em `LocalFile` para exercitar o fluxo de
        # upload pequeno.
        local_file = LocalFile.from_path(r'/Users/eu.phietro/Developer/eng_software/programacao_web/meu_portifolio/assets/Currículo - Henry Aguiar.pdf')

        # 11. Envia o arquivo para a pasta remota garantida acima.
        created_file = await sharepoint.upload_small_file(
            default_drive,    
            created_hierarcly,    
            local_file, 
            'rename',   
        )
        print(created_file)
        
    except Exception as error:
        raise error

if __name__ == '__main__':
    asyncio.run(main())
