# MS Cronos

MS Cronos e um projeto Python leve para automatizar o envio de arquivos locais para bibliotecas de documentos do SharePoint usando Microsoft Graph.

O objetivo inicial e construir um core pequeno em torno do `msgraph-sdk`, com uma interface mais simples para resolver sites, encontrar drives, listar pastas e enviar arquivos. Em uma segunda etapa, esse core deve servir como base para um daemon no Windows. Depois disso, uma API simples podera criar jobs de execucao para o daemon processar.

## Objetivo

Criar um wrapper Python para SharePoint que esconda a parte verbosa do Microsoft Graph SDK e ofereca operacoes diretas como:

- resolver um site por URL do SharePoint;
- obter a biblioteca de documentos padrao;
- listar itens de uma pasta;
- encontrar uma pasta pelo nome;
- criar pastas;
- enviar arquivos locais para uma pasta no SharePoint.

## Escopo Atual

Neste momento o projeto esta em fase de descoberta e construcao do core. O arquivo `main.py` funciona como laboratorio para validar chamadas reais contra o Microsoft Graph.

Fluxo ja validado:

1. Autenticacao com `ClientSecretCredential`.
2. Criacao do `GraphServiceClient`.
3. Resolucao de um site SharePoint.
4. Obtencao do drive padrao do site.
5. Obtencao do item `root` do drive.
6. Listagem dos filhos de uma pasta.
7. Criacao de pasta com `children.post(...)`.

## Arquitetura Planejada

O projeto deve evoluir em camadas pequenas:

```text
ms_cronos/
  main.py                  # laboratorio atual
  ms_cronos/
    graph_client.py         # cria e fecha o GraphServiceClient
    sharepoint.py           # wrapper principal de SharePoint/Drive
    models.py               # dataclasses leves para configs e resultados
    jobs.py                 # contratos de jobs futuros
  README.md
  requirements.txt
```

O core nao deve depender de API web, scheduler ou servico Windows. Essas partes entram depois.

## Core Desejado

A primeira versao do wrapper deve oferecer uma classe parecida com:

```python
class SharePointDriveClient:
    async def resolve_site(self, sharepoint_url: str): ...
    async def get_default_drive(self, site_id: str): ...
    async def list_children(self, drive_id: str, folder_id: str): ...
    async def find_child_folder(self, drive_id: str, parent_id: str, name: str): ...
    async def create_folder(self, drive_id: str, parent_id: str, name: str): ...
    async def upload_file(self, drive_id: str, parent_id: str, local_path: str): ...
```

A ideia e que o restante do sistema nao precise conhecer detalhes como:

- `client.sites.with_url(...)`;
- `client.drives.by_drive_id(...).items.by_drive_item_id(...).children.get()`;
- `DriveItem`;
- `Folder`;
- URL especial de upload `items/{parent-id}:/{filename}:/content`.

## Microsoft Graph

A URL base usada pelo SDK e:

```text
https://graph.microsoft.com/v1.0
```

Rotas importantes para este projeto:

```http
GET /sites/{hostname}:/{site-path}
GET /sites/{site-id}/drive
GET /drives/{drive-id}/root
GET /drives/{drive-id}/items/{folder-id}/children
POST /drives/{drive-id}/items/{folder-id}/children
PUT /drives/{drive-id}/items/{folder-id}:/{filename}:/content
PUT /drives/{drive-id}/items/{file-id}/content
```

Conceitos principais:

- `Site`: representa um site do SharePoint.
- `Drive`: representa uma biblioteca de documentos.
- `DriveItem`: representa uma pasta ou arquivo.
- `children`: lista itens dentro de uma pasta.
- `content`: conteudo binario de um arquivo.

## Autenticacao

O projeto usa autenticacao de aplicacao com `ClientSecretCredential`.

Variaveis esperadas no `.env`:

```env
CLIENT_ID=
CLIENT_SECRET=
CLIENT_TENANT=
```

O `.env` nao deve ser versionado nem compartilhado.

Permissoes esperadas no Microsoft Graph dependem do escopo final, mas para o core de arquivos normalmente serao necessarias permissoes de aplicacao relacionadas a arquivos e sites, como `Sites.ReadWrite.All` ou uma configuracao mais restrita com `Sites.Selected`.

## Instalacao

```bash
python -m venv venv
source venv/bin/activate
pip install -r requirements.txt
```

No Windows:

```powershell
python -m venv venv
.\venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

## Uso Atual

O ponto de entrada atual e:

```bash
python main.py
```

Esse script ainda e experimental. Ele deve ser usado para aprender e validar o comportamento do SDK antes de mover o codigo para o wrapper.

## Upload de Arquivos

Para criar uma pasta, o SDK permite usar objetos:

```python
from msgraph.generated.models.drive_item import DriveItem
from msgraph.generated.models.folder import Folder

folder = DriveItem(
    name="Nova Pasta",
    folder=Folder(),
    additional_data={
        "@microsoft.graph.conflictBehavior": "rename",
    },
)

created = await client.drives.by_drive_id(drive_id) \
    .items.by_drive_item_id(parent_folder_id) \
    .children.post(folder)
```

Para enviar arquivo novo para uma pasta, a rota mais simples do Graph usa o nome do arquivo no caminho:

```http
PUT /drives/{drive-id}/items/{folder-id}:/{filename}:/content
```

O wrapper deve encapsular essa URL para que o restante do projeto use apenas:

```python
await sharepoint.upload_file(
    drive_id=drive_id,
    parent_id=folder_id,
    local_path="C:/arquivos/curriculo.pdf",
)
```

Para sobrescrever um arquivo existente, o SDK consegue usar o objeto do arquivo diretamente:

```python
updated = await client.drives.by_drive_id(drive_id) \
    .items.by_drive_item_id(file_id) \
    .content.put(file_bytes)
```

## Daemon Windows

O daemon Windows sera uma camada posterior. A responsabilidade dele deve ser:

- carregar configuracoes;
- buscar jobs pendentes;
- executar upload de arquivos;
- registrar resultado da execucao;
- tentar novamente em falhas recuperaveis.

O daemon nao deve conter logica direta do Microsoft Graph. Ele deve chamar o wrapper do core.

## API de Jobs

A API futura deve ser pequena e focada em cadastrar execucoes.

Um job minimo pode conter:

```json
{
  "sharepoint_url": "https://tenant.sharepoint.com/sites/RHConecta",
  "drive": "default",
  "target_folder": "Curriculos",
  "local_path": "C:/arquivos/curriculo.pdf",
  "conflict_behavior": "rename"
}
```

Estados iniciais sugeridos:

- `pending`
- `running`
- `done`
- `failed`

## Principios do Projeto

- O core deve ser pequeno e facil de testar.
- O wrapper deve esconder a verbosidade do SDK.
- O daemon deve apenas orquestrar jobs.
- A API deve apenas criar e consultar jobs.
- Logs e erros precisam ser legiveis para facilitar operacao no Windows.

## Proximos Passos

1. Extrair a autenticacao de `main.py` para `graph_client.py`.
2. Criar `SharePointDriveClient`.
3. Implementar `resolve_site`.
4. Implementar `get_default_drive`.
5. Implementar `list_children`.
6. Implementar `find_child_folder`.
7. Implementar `create_folder`.
8. Implementar `upload_file`.
9. Criar um teste manual de upload pequeno.
10. Separar o laboratorio `main.py` do core reutilizavel.
