# MS Cronos

MS Cronos e um projeto Python para navegacao e upload de conteudo local para
bibliotecas de documentos do SharePoint via Microsoft Graph.

O foco atual esta no `core/`: uma camada pequena, tipada e semantica que
esconde a verbosidade do SDK oficial e expoe operacoes previsiveis sobre site,
drive, pasta e arquivo.

## Objetivo

O objetivo final do projeto e permitir upload de diretorios locais complexos
para o SharePoint, preservando a estrutura de pastas remotas e reduzindo o
contato do restante do sistema com detalhes do Graph.

Fluxo alvo:

```text
diretorio local
  -> resolver site
  -> escolher drive
  -> obter root
  -> encontrar ou criar pastas remotas
  -> enviar arquivos pequenos e grandes
  -> retornar resultados semanticos
```

## Estado Atual

O Core ja implementa:

- autenticacao com `ClientSecretCredential`;
- construcao de `GraphServiceClient`;
- resolucao de site por URL humana do SharePoint;
- listagem de drives de um site;
- obtencao do drive padrao;
- obtencao de drive especifico por id;
- obtencao do item raiz do drive;
- listagem de filhos imediatos de uma pasta remota;
- busca imediata de filhos por nome, id e `web_url`;
- busca tipada de pasta filha e arquivo filho;
- criacao de pasta remota;
- garantia incremental de caminho remoto;
- upload pequeno com `conflict_behavior`;
- parse de models do SDK para models internos;
- traducao inicial de `ODataError` para erros semanticos do Core.

Ainda faltam principalmente:

- upload de arquivos grandes com upload session;
- leitura e mapeamento de arvores locais;
- upload completo de diretorios locais complexos.

## Estrutura do Repositorio

```text
ms_cronos/
├─ core/
│  ├─ __init__.py
│  ├─ builders.py
│  ├─ errors.py
│  ├─ graph_client.py
│  ├─ models.py
│  ├─ parse.py
│  ├─ sharepoint.py
│  ├─ urls.py
│  └─ utils.py
├─ docs/
│  ├─ LLM.md
│  └─ SPEC.md
├─ main.py
├─ README.md
└─ requirements.txt
```

## Mapa Conceitual

Hierarquia remota importante:

```text
Tenant
  -> Site
  -> Drive (document library)
  -> DriveItem
     -> Folder
     -> File
```

No Core isso aparece como:

```text
URL humana
  -> SiteRef
  -> DriveRef
  -> DriveItemRef
```

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

## Configuracao

O projeto usa autenticacao de aplicacao com `ClientSecretCredential`.

Variaveis esperadas no `.env`:

```env
CLIENT_ID=
CLIENT_SECRET=
CLIENT_TENANT=
```

## Exemplo Rapido

Exemplo minimo de resolucao de site, navegacao e upload pequeno:

```python
import asyncio
from os import getenv

from core import GraphClientManager, GraphCredentials, LocalFile, SharePointService


async def main() -> None:
    credentials = GraphCredentials(
        client_id=getenv("CLIENT_ID", ""),
        client_secret=getenv("CLIENT_SECRET", ""),
        tenant_id=getenv("CLIENT_TENANT", ""),
    )

    site_url = "https://tenant.sharepoint.com/sites/RHConecta"

    async with GraphClientManager(credentials) as manager:
        sharepoint = SharePointService(manager)

        site = await sharepoint.resolve_site(site_url)
        drive = await sharepoint.get_default_drive(site)
        root = await sharepoint.get_drive_root(drive)

        target_folder = await sharepoint.ensure_remote_folder_path(
            drive,
            root,
            ("datasets", "2026", "06", "04"),
        )

        local_file = LocalFile.from_path("/tmp/relatorio.csv")

        result = await sharepoint.upload_small_file(
            drive,
            target_folder,
            local_file,
            conflict_behavior="rename",
        )

        print(site)
        print(drive)
        print(target_folder)
        print(result)


if __name__ == "__main__":
    asyncio.run(main())
```

## Exemplos de Uso

### Resolver um site

```python
site = await sharepoint.resolve_site(
    "https://tenant.sharepoint.com/sites/RHConecta"
)

print(site.id)
print(site.web_url)
```

### Listar drives do site

```python
drives = await sharepoint.list_site_drives(site)

for drive in drives:
    print(drive.id, drive.name, drive.drive_type)
```

### Obter o drive padrao e a raiz

```python
drive = await sharepoint.get_default_drive(site)
root = await sharepoint.get_drive_root(drive)

print(drive.name)
print(root.name, root.is_folder)
```

### Listar filhos imediatos

```python
children = await sharepoint.list_children(drive, root)

for child in children:
    print(child.name, child.is_folder, child.is_file)
```

Importante: `list_children(...)` nao e recursivo. Cada chamada representa um
unico nivel da arvore remota.

### Encontrar uma pasta filha

```python
curriculos = await sharepoint.find_child_folder(drive, root, "Curriculos")

if curriculos is None:
    print("Pasta nao encontrada")
else:
    print(curriculos.id, curriculos.web_url)
```

### Criar uma pasta filha

```python
created = await sharepoint.create_folder(
    drive,
    root,
    "Pasta_de_Teste",
    conflict_behavior="rename",
)

print(created)
```

### Garantir um caminho remoto inteiro

```python
leaf = await sharepoint.ensure_remote_folder_path(
    drive,
    root,
    ("datasets", "financeiro", "2026", "06"),
    conflict_behavior="fail",
)

print(leaf.name)
```

### Upload pequeno

```python
local_file = LocalFile.from_path("/tmp/relatorio.pdf")

result = await sharepoint.upload_small_file(
    drive,
    root,
    local_file,
    conflict_behavior="replace",
)

print(result.remote_name)
print(result.item.id)
```

`conflict_behavior` hoje aceita:

- `fail`
- `rename`
- `replace`

O fluxo de upload pequeno atualmente rejeita arquivos acima de `250_000_000`
bytes.

## Colecoes do Core

O Core evita retornar listas cruas quando existe uma colecao semantica mais
expressiva.

Exemplo:

```python
children = await sharepoint.list_children(drive, root)

print(children.counter)
print(children.is_empty)
print(children.first())
print(children.to_list())
```

## Modulos Principais

### `core/models.py`

Contem:

- modelos semanticos como `SiteRef`, `DriveRef`, `DriveItemRef`, `LocalFile`;
- colecoes genericas como `Collection_`, `FrozenCollection`, `MutableCollection`;
- colecoes concretas como `SiteRefCollection`, `DriveRefCollection`,
  `DriveItemCollection` e `LocalFileCollection`;
- contratos de upload como `StagingContentUpload` e `UploadResult`.

### `core/graph_client.py`

Responsavel por:

- validar credenciais;
- criar o `GraphServiceClient`;
- encapsular o ciclo de vida do `ClientSecretCredential`;
- permitir uso com `async with`.

### `core/parse.py`

Responsavel por:

- converter `Site`, `Drive` e `DriveItem` do SDK em contratos internos;
- converter collection responses do SDK em colecoes concretas;
- adaptar `ODataError` para a hierarquia de erros do Core;
- adaptar caminhos locais para `LocalFile`.

### `core/builders.py`

Responsavel por:

- montar o `DriveItem` usado para criacao de pastas;
- preparar o staging de upload pequeno;
- validar nomes remotos e `conflict_behavior`.

### `core/sharepoint.py`

Contem o servico principal. Hoje ele expoe:

```python
class SharePointService:
    async def resolve_site(self, sharepoint_url: str) -> SiteRef: ...
    async def list_site_drives(self, site_ref: SiteRef) -> DriveRefCollection: ...
    async def get_default_drive(self, site_ref: SiteRef) -> DriveRef: ...
    async def get_drive(self, drive_ref: DriveRef) -> DriveRef: ...
    async def get_drive_root(self, drive_ref: DriveRef) -> DriveItemRef: ...
    async def list_children(
        self,
        drive_ref: DriveRef,
        parent_item_ref: DriveItemRef,
    ) -> DriveItemCollection: ...
    async def find_child_by_name(
        self,
        drive_ref: DriveRef,
        parent_item_ref: DriveItemRef,
        name: str,
    ) -> DriveItemRef | None: ...
    async def find_child_by_id(
        self,
        drive_ref: DriveRef,
        parent_item_ref: DriveItemRef,
        drive_id: str,
    ) -> DriveItemRef | None: ...
    async def find_child_by_web_url(
        self,
        drive_ref: DriveRef,
        parent_item_ref: DriveItemRef,
        web_url: str,
    ) -> DriveItemRef | None: ...
    async def find_child_folder(
        self,
        drive_ref: DriveRef,
        parent_item_ref: DriveItemRef,
        name: str,
    ) -> DriveItemRef | None: ...
    async def find_child_file(
        self,
        drive_ref: DriveRef,
        parent_item_ref: DriveItemRef,
        name: str,
    ) -> DriveItemRef | None: ...
    async def find_child_folder_by_id(
        self,
        drive_ref: DriveRef,
        parent_item_ref: DriveItemRef,
        drive_id: str,
    ) -> DriveItemRef | None: ...
    async def find_child_folder_web_url(
        self,
        drive_ref: DriveRef,
        parent_item_ref: DriveItemRef,
        web_url: str,
    ) -> DriveItemRef | None: ...
    async def create_folder(
        self,
        drive_ref: DriveRef,
        parent_item_ref: DriveItemRef,
        folder_name: str,
        conflict_behavior: str = "fail",
    ) -> DriveItemRef: ...
    async def ensure_remote_folder_path(
        self,
        drive_ref: DriveRef,
        root_item: DriveItemRef,
        folders_parts: tuple[str, ...],
        conflict_behavior: str = "fail",
    ) -> DriveItemRef: ...
    async def upload_small_file(
        self,
        drive_ref: DriveRef,
        parent_item_ref: DriveItemRef,
        local_file: LocalFile,
        conflict_behavior: str = "fail",
    ) -> UploadResult: ...
```

## Rotas Mais Importantes do Graph

```http
GET /sites/{hostname}:/{site-path}
GET /sites/{site-id}/drives
GET /sites/{site-id}/drive
GET /drives/{drive-id}
GET /drives/{drive-id}/root
GET /drives/{drive-id}/items/{item-id}/children
POST /drives/{drive-id}/items/{item-id}/children
PUT /drives/{drive-id}/items/{item-id}:/{filename}:/content
PUT /drives/{drive-id}/items/{item-id}/content
```

## Uso Atual do Repositorio

O arquivo `main.py` ainda funciona como laboratorio manual e roteiro de
validacao do fluxo real contra o Graph:

```bash
python main.py
```

## Principios do Projeto

- o Core deve ser pequeno e testavel;
- o SharePoint deve ser exposto por contratos semanticos, nao por envelopes do SDK;
- o parser deve concentrar adaptacao de models e erros;
- a camada de servico deve orquestrar chamadas, nao inventar estruturas paralelas;
- a evolucao futura para daemon e jobs deve reutilizar o Core sem acoplamento.

## Leitura Complementar

- [docs/SPEC.md](/Users/eu.phietro/Developer/plangecon/ms_cronos/docs/SPEC.md)
- [docs/LLM.md](/Users/eu.phietro/Developer/plangecon/ms_cronos/docs/LLM.md)
