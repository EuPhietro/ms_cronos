import asyncio # Módulo para operações assincronas
import logging # Módulo para logging 
from urllib.parse import urlparse
from os import getenv # Módulo para interação com o sistema operacional

 
from azure.identity.aio import ClientSecretCredential # Classe basica para Credenciais
from msgraph.graph_service_client import GraphServiceClient # Client para a Graph API
from msgraph.generated.models.drive_item import DriveItem
from msgraph.generated.models.folder import Folder
from msgraph.generated.models.o_data_errors.o_data_error import ODataError # Erro lançado OData
from dotenv import load_dotenv # Importando a função load_dotenv
from rich import print
from rich.json import JSON
from rich.traceback import install



# Carregando variáveis de ambiente
load_dotenv()
install(show_locals=True)



def build_graph_site_url(sharepoint_url: str) -> str:
    '''
    Retorna uma 
    '''
    # Type: ParseResult, que contem: username, password, hostname, path e uris
    # Faz o parse da URL um objeto PaserResult com os componentes da url como atributos
    parsed = urlparse(sharepoint_url)
    
    if not parsed: 
        raise ValueError('')
    
    # Pegamos o host_name da URL
    hostname = parsed.hostname
    # E por último o caminho removendo o '/'ao final
    path = parsed.path.rstrip('/') # Remove caracteres de '/' localizados à direita 

    # Verificamos se ele é Nulo, se for, levanta um erro
    if not  hostname:
        raise ValueError('URL do SharePoint inválida."')
    # Verificamos se path é Nulo, se for, retorna apenas o caminho base
    if not path:
        return f'https://graph.microsoft.com/v1.0/sites/{hostname}'
    # Se não, retorna o caminho completo com o complemento
    return f'https://graph.microsoft.com/v1.0/sites/{hostname}:{path}'


# A função main é assincrona
async def main() -> None:
    client_id = getenv('CLIENT_ID', '') # Carregano client_id das variáveis de ambiente
    client_secret = getenv('CLIENT_SECRET', '') # Carregando secret das variáveis de ambiente
    tenant_id = getenv('CLIENT_TENANT', '') # Carregando ID do TENANT (Locatório)
    
    
    # Instancia um objeto do tipo ClientSecretCredential que resolve a autenticação
    credential = ClientSecretCredential(
        tenant_id=tenant_id, 
        client_id=client_id, 
        client_secret=client_secret
    ) # Instanciando ClientSecrets
    
    # O escopo por padrão é a url padrão da API
    scopes = ['https://graph.microsoft.com/.default'] # URL para o Gateway padrão da Microsoft
    # Aqui nós inicializamos uma instância do GraphServiceClient que é um client para API
    client = GraphServiceClient(credential, scopes=scopes) # Inicializando o GraphServiceClient
    
    
    # Executamos a operação dentro de um try except
    try:
        print("\n[INFO] Acessando o Microsoft Graph e buscando a coleção de sites...")
        
        # 2. Mudança de abordagem: Em vez de consultar um ID específico, 
        # vamos pedir para o Graph listar os sites da raiz do tenant.
        # Isso testa se a permissão da aplicação está funcionando na raiz do SharePoint.
        sharepoint_site = build_graph_site_url('https://plangeconcombr.sharepoint.com/sites/RHConecta')
        
        site = await client.sites.with_url(sharepoint_site).get()
        
        if not site:
            raise ODataError
        
        site_id = site.additional_data.get('id', '')
        drive = await client.sites.by_site_id(site_id).drive.get()
        

        if not drive:
            raise ODataError
        
        print(f"Acessando o drive: {drive.name} | {drive.id}")
        
        root = await client.drives.by_drive_id(drive.id).root.get()
        children = await client.drives.by_drive_id(drive.id).items.by_drive_item_id(root.id).children.get()
        
        if children and children.value:
            for item in children.value:
                item_type = "pasta" if item.folder else "arquivo"
                print(f"{item.name} | {item.id} | {item_type}")
                
      
        target = None  
        
        if  children and children.value:
           for drive_item in children.value:
               if drive_item.name == 'datasets':
                   target = drive_item
                   
        
        datasets = await client.drives.by_drive_id(drive.id).items.by_drive_item_id(target.id).children.get()
        
        print(datasets)
        
        new_folder = DriveItem(
            name='Folder de Teste2',
            folder=Folder(),
            additional_data= {
            "@microsoft.graph.conflictBehavior": "rename"
            },
        )
        
        created = await client.drives.by_drive_id(drive.id).items.by_drive_item_id(target.id).children.post(new_folder)

        print(created)
        
            
    except ODataError as error:
        print("\n❌ Erro detalhado do Microsoft Graph:")
        if error.error:
            print(f"Código do Erro: {error.error.code}")
            print(f"Mensagem: {error.error.message}")
        else:
            print(f"ODataError retornado: {error}")
        raise 
    except Exception as error:
        print(f'\n❌ Falha inesperada ocorreu: {error} ({error.__class__.__name__})')
    finally:
        await credential.close()

if __name__ == '__main__':
    asyncio.run(main())