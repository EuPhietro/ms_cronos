import asyncio # Módulo para operações assincronas
from os import getenv # Módulo para interação com o sistema operacional

 
from core import GraphCredentials, GraphClientManager, SharePointService

from dotenv import load_dotenv # Importando a função load_dotenv
from rich import print
from rich.json import JSON
from rich.traceback import install



# Carregando variáveis de ambiente
load_dotenv()
install(show_locals=True)






# A função main é assincrona
async def main() -> None:
    client_id = getenv('CLIENT_ID', '') # Carregano client_id das variáveis de ambiente
    client_secret = getenv('CLIENT_SECRET', '') # Carregando secret das variáveis de ambiente
    tenant_id = getenv('CLIENT_TENANT', '') # Carregando ID do TENANT (Locatório)
    
    
    credentials = GraphCredentials(client_id,client_secret,tenant_id) # Classe de credentials interno
    site_url= 'https://plangeconcombr.sharepoint.com/sites/RHConecta' # Url padrão
    graph_client_manager = GraphClientManager(credentials) # Instanciando um GraphClientManager
    sharepoint = SharePointService(graph_client_manager) # Instanciando um objeto do tipo sharepoint_service
    try:
        sites =  await sharepoint.resolve_site(site_url) # Obtendo o site especificado
        libraries = await sharepoint.list_site_drives(sites) # Obtendo a biblioteca de documentos do site especificado
        root_drive = await sharepoint.get_default_drive(sites)
        other_drive = await sharepoint.get_drive(root_drive)
        drive_root = await sharepoint.get_drive_root(other_drive)
        childrens = await sharepoint.list_children(other_drive, drive_root)
        
        print(libraries) # Visualizando o conteúdo
        print(root_drive)
        print(other_drive)
        print(drive_root)
        print(childrens)
    except Exception as error: 
        raise error

if __name__ == '__main__':
    asyncio.run(main())