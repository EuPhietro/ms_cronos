"""Laboratorio manual para validar o Core contra o Microsoft Graph.

Este arquivo nao faz parte da API do Core. Ele serve como roteiro executavel para
testar manualmente credenciais, resolucao de site, navegacao por drive e leitura
da raiz enquanto a implementacao evolui.
"""

import asyncio
from logging import root
from os import getenv

from dotenv import load_dotenv
from rich import print
from rich.traceback import install

from core import GraphClientManager, GraphCredentials, SharePointService, LocalFile


# Laboratorio manual para exercitar o Core contra o Microsoft Graph.
load_dotenv()
install(show_locals=True)


async def main() -> None:
    # Credenciais de aplicacao carregadas do ambiente local.
    client_id = getenv('CLIENT_ID', '')
    client_secret = getenv('CLIENT_SECRET', '')
    tenant_id = getenv('CLIENT_TENANT', '')

    credentials = GraphCredentials(client_id, client_secret, tenant_id)
    site_url = 'https://plangeconcombr.sharepoint.com/sites/RHConecta'
    graph_client_manager = GraphClientManager(credentials)
    sharepoint = SharePointService(graph_client_manager)
    try:
        # Fluxo validado: site -> drives -> drive padrao -> root do drive.
        site = await sharepoint.resolve_site(site_url)
        print(site)
        drives_library = await sharepoint.list_site_drives(site)
        print(drives_library)
        default_drive = await sharepoint.get_default_drive(site)
        print(default_drive)
        root_item = await sharepoint.get_drive_root(default_drive)
        print(root_item)
        children = await sharepoint.list_children(default_drive, root_item)
        print(children)
        curriculos_drive_item = children[0]
        print(curriculos_drive_item)
        curriculos_items = await sharepoint.list_children(default_drive, curriculos_drive_item)
        print(curriculos_items)
        created_drive_item = await sharepoint.create_folder(default_drive, curriculos_drive_item, 'Pasta_de_Teste_1345',conflict_behavior='rename')
        print(created_drive_item)
        created_hierarcly = await sharepoint.ensure_remote_folder_path(
            default_drive,
            root_item,
            folders_parts=('datasets2','2026','05','21','20-00','1')
        )
        print(created_hierarcly)
        local_file = LocalFile.from_path(r'/Users/eu.phietro/Developer/eng_software/programacao_web/meu_portifolio/assets/Currículo - Henry Aguiar.pdf')
        created_file = await sharepoint.upload_small_file(default_drive, root_item,local_file,'rename')
        print(created_file)
    except Exception as error:
        raise error

if __name__ == '__main__':
    asyncio.run(main())
