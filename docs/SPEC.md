# MS Cronos Specification

Este documento descreve a especificacao funcional do estado atual do projeto e
do que ainda falta para completar o fluxo de upload de diretorios locais
complexos.

## Visao Geral

MS Cronos deve automatizar navegacao e upload de conteudo local para
SharePoint usando Microsoft Graph.

O Core do projeto foi desenhado para funcionar como um wrapper pequeno e
semantico sobre o SDK oficial, expondo contratos internos previsiveis em vez de
envelopes crus do Graph.

Fluxo de produto desejado:

```text
diretorio local
  -> leitura da arvore local
  -> resolve site
  -> escolhe drive
  -> encontra ou cria pastas remotas
  -> envia arquivos pequenos ou grandes
  -> retorna resultados semanticos
```

## Objetivos

- encapsular a verbosidade do Microsoft Graph;
- expor contratos internos pequenos e tipados;
- suportar navegacao semantica por site, drive e item;
- suportar criacao incremental de pastas remotas;
- suportar upload pequeno com politica de conflito clara;
- suportar upload grande com sessao resumivel;
- preparar a abstracao completa do SDK e o upload recursivo de diretorios.

## Nao Objetivos da Fase Atual

Na fase atual, o projeto ainda nao busca:

- sincronizacao bidirecional;
- API HTTP publica;
- UI;
- monitoramento automatico de pastas locais;
- suporte generico a todos os recursos do Microsoft Graph;
- daemon final de producao.

## Camadas Desejadas

Arquitetura alvo:

```text
Presentation / API
  -> recebe comandos externos

Application / Jobs
  -> orquestra fluxos de negocio

Daemon / Worker
  -> executa jobs em background

Core
  -> implementa SharePoint e Graph

Infrastructure
  -> ambiente, logging, persistencia, configuracao
```

A implementacao atual esta concentrada no `Core`.

## Estrutura Atual

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

## Hierarquia Conceitual

Para este projeto, a estrutura remota relevante e:

```text
Tenant
  -> Site
  -> Drive
  -> DriveItem
     -> Folder
     -> File
```

Dentro do Core, os contratos equivalentes sao:

```text
SiteRef
DriveRef
DriveItemRef
LocalFile
UploadFileResult
```

## Contratos do Core

### Modelos

Os modelos internos mais importantes hoje sao:

- `GraphCredentials`
- `SiteRef`
- `DriveRef`
- `DriveItemRef`
- `LocalFile`
- `StagingContentUpload`
- `UploadFileResult`

### Exemplo de modelos

```python
from pathlib import Path

from core import DriveItemRef, DriveRef, GraphCredentials, LocalFile, SiteRef

credentials = GraphCredentials(
    client_id="app-id",
    client_secret="secret",
    tenant_id="tenant-id",
)

site = SiteRef(
    id="tenant.sharepoint.com,site-guid,web-guid",
    name="RHConecta",
    display_name="RH Conecta",
    web_url="https://tenant.sharepoint.com/sites/RHConecta",
)

drive = DriveRef(
    id="b!abc123",
    name="Documents",
    web_url="https://tenant.sharepoint.com/sites/RHConecta/Shared%20Documents",
    drive_type="documentLibrary",
)

root = DriveItemRef(
    id="01ABCDEF",
    name="root",
    web_url="https://tenant.sharepoint.com/sites/RHConecta/Shared%20Documents",
    is_folder=True,
    is_file=False,
    size=0,
)

local_file = LocalFile.from_uri(Path("/tmp/relatorio.csv"))
```

### Colecoes

O projeto usa colecoes semanticas em vez de listas cruas sempre que isso ajuda
na legibilidade do dominio.

Colecoes mais importantes:

- `Collection_[T]`
- `FrozenCollection[T]`
- `MutableCollection[T]`
- `SiteRefCollection`
- `DriveRefCollection`
- `DriveItemCollection`
- `LocalFileCollection`

### Exemplo de uso das colecoes

```python
children = await sharepoint.list_children(drive, root)

print(children.counter)
print(children.is_empty)
print(children.first())

for child in children:
    print(child.name)
```

## Modulos do Core

### `core/graph_client.py`

Responsavel por:

- validar credenciais;
- criar `GraphServiceClient`;
- encapsular o ciclo de vida do `ClientSecretCredential`;
- permitir uso com `async with`.

Exemplo:

```python
from core import GraphClientManager, GraphCredentials

credentials = GraphCredentials(
    client_id="app-id",
    client_secret="secret",
    tenant_id="tenant-id",
)

async with GraphClientManager(credentials) as manager:
    client = manager.client
```

### `core/models.py`

Responsavel por:

- contratos semanticos do Core;
- colecoes reutilizaveis;
- separacao entre leitura, imutabilidade e mutabilidade;
- contratos de staging e resultado de upload.

### `core/parse.py`

Responsavel por:

- converter `Site`, `Drive` e `DriveItem` do SDK em modelos internos;
- converter collection responses do SDK em colecoes concretas;
- converter `Path` local em `LocalFile`;
- traduzir `ODataError` para a hierarquia de erros do Core.

Exemplo de uso interno:

```python
from core import parse_drive_item

drive_item_ref = parse_drive_item(graph_drive_item)
```

### `core/urls.py`

Responsavel por:

- validar URLs humanas do SharePoint;
- montar a rota de resolucao de site para o Graph;
- montar fragmentos e URLs do fluxo de upload pequeno.

Exemplo:

```python
from core import build_graph_site_url

graph_url = build_graph_site_url("https://tenant.sharepoint.com/sites/RHConecta")
```

### `core/builders.py`

Responsavel por:

- montar o `DriveItem` de criacao de pasta;
- montar o `StagingContentUpload` do upload pequeno;
- validar nome remoto e `conflict_behavior`.

Exemplo:

```python
from core import LocalFile
from core.builders import build_folder_drive_item, build_upload_content

body = build_folder_drive_item("Relatorios", conflict_behavior="rename")

local_file = LocalFile.from_uri("/tmp/relatorio.pdf")
staging = build_upload_content(local_file, None, conflict_behavior="replace")
```

### `core/sharepoint.py`

Servico principal do Core.

Responsabilidades:

- chamar o Graph;
- validar regras de dominio;
- delegar parse para `core.parse`;
- devolver apenas contratos internos.

## API Atual do Servico

Hoje o `SharePointService` expoe aproximadamente:

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
    ) -> UploadFileResult: ...
    async def upload_large_file(
        self,
        drive_ref: DriveRef,
        parent_item_ref: DriveItemRef,
        local_file: LocalFile,
        conflict_behavior: str = "fail",
    ) -> UploadResult[DriveItem]: ...
```

## Exemplo Completo de Uso

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

    async with GraphClientManager(credentials) as manager:
        sharepoint = SharePointService(manager)

        site = await sharepoint.resolve_site(
            "https://tenant.sharepoint.com/sites/RHConecta"
        )
        drive = await sharepoint.get_default_drive(site)
        root = await sharepoint.get_drive_root(drive)

        children = await sharepoint.list_children(drive, root)
        print(children.counter)

        target = await sharepoint.ensure_remote_folder_path(
            drive,
            root,
            ("datasets", "2026", "06", "04"),
        )

        local_file = LocalFile.from_uri("/tmp/relatorio.csv")
        result = await sharepoint.upload_small_file(
            drive,
            target,
            local_file,
            conflict_behavior="rename",
        )

        print(result.remote_name)
        print(result.item.web_url)


asyncio.run(main())
```

## Comportamentos Importantes

### Navegacao

- `list_children(...)` nao e recursivo;
- buscas por nome, id e `web_url` operam apenas sobre filhos imediatos;
- ids de `DriveItem` sao usados dentro do contexto do `drive_ref`.

### Criacao de pasta

- `create_folder(...)` exige que o item pai seja pasta;
- o body do Graph e montado em `core/builders.py`;
- `conflict_behavior` aceito: `fail`, `rename`, `replace`.

### Upload pequeno

- o fluxo trabalha com o arquivo inteiro em memoria;
- o limite atual e `250_000_000` bytes;
- se o arquivo remoto nao existir, o Core cria por nome;
- se existir:
  - `fail` levanta erro;
  - `rename` cria com novo nome;
  - `replace` substitui o conteudo do item existente.

### Upload grande

O fluxo atual:

1. valida o item pai e o arquivo local;
2. cria uma `UploadSession` por path remoto;
3. abre o arquivo como stream binario;
4. usa `LargeFileUploadTask` para enviar chunks;
5. devolve `UploadResult[DriveItem]` do SDK.

Exemplo atual:

```python
local_file = LocalFile.from_uri("/tmp/backup.zip")

sdk_result = await sharepoint.upload_large_file(
    drive,
    root,
    local_file,
    conflict_behavior="rename",
)
```

O quinto passo deve mudar na fase 2: o retorno publico esperado passa a ser
`UploadFileResult`, contendo um `DriveItemRef`.

## Especificacao da Fase 2

A fase 2 deve criar uma fronteira publica independente do SDK.

### Obrigacoes

- nenhum metodo publico deve exigir um model gerado pelo Graph;
- nenhum metodo publico deve retornar um model do Graph ou do `msgraph-core`;
- request bodies gerados devem ser construidos em builders internos;
- respostas do SDK devem passar por parsers ou factories;
- excecoes externas devem ser convertidas para erros semanticos;
- wrappers de comportamento devem esconder tasks e request adapters;
- tipos compartilhados, como conflito e resultado, devem ter uma unica
  definicao no Core.

### Fluxo desejado do upload grande

```text
LocalFile + DriveRef + DriveItemRef
  -> LargeFileUploadRequest
  -> wrapper/executor interno
  -> UploadSession do SDK
  -> LargeFileUploadTask do SDK
  -> UploadResult[DriveItem] do SDK
  -> factory/parser
  -> UploadFileResult
```

### Exemplo publico desejado

```python
result = await sharepoint.upload_large_file(
    drive,
    root,
    LocalFile.from_uri("/tmp/backup.zip"),
    conflict_behavior="rename",
)

print(result.item.id)
print(result.remote_name)
print(result.source_path)
```

## Erros Semanticos

O Core evita vazar excecoes cruas do SDK.

Exemplos relevantes:

- `SiteResolutionError`
- `DriveNotFoundError`
- `DriveItemNotFoundError`
- `FolderNotFoundError`
- `NotAFolderError`
- `NotAFileError`
- `InvalidRemoteNameError`
- `InvalidConflictBehaviorError`
- `FileAlreadyExistError`
- `FileVeryLargeError`
- `SmallFileUploadError`
- `GraphPermissionError`
- `GraphAuthenticationError`
- `GraphRequestError`
- `GraphResourceConflictError`

## ODataError

O parse de `ODataError` ja cobre, entre outros:

- `accessDenied` -> `GraphPermissionError`
- `authorizationRequestDenied` -> `GraphPermissionError`
- `unauthenticated` -> `GraphAuthenticationError`
- `invalidAuthenticationToken` -> `GraphAuthenticationError`
- `invalidRequest` -> `GraphRequestError`
- `badRequest` -> `GraphRequestError`
- `badArgument` -> `GraphRequestError`
- `nameAlreadyExists` -> `GraphResourceConflictError`
- `conflict` -> `GraphResourceConflictError`
- `itemNotFound` -> erro contextual conforme a operacao
- `resourceNotFound` -> erro contextual conforme a operacao

## Rotas Principais do Graph

Rotas mais importantes para o estado atual:

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
POST /drives/{drive-id}/items/{parent-id}:/{filename}:/createUploadSession
```

## Proximos Passos

As implementacoes de maior prioridade agora sao:

1. definir erros semanticos do upload grande;
2. criar wrappers para sessao e executor de chunks;
3. criar factory/parser de resultado para `UploadFileResult`;
4. remover tipos do SDK das assinaturas publicas;
5. ler e mapear diretorios locais;
6. enviar arvores locais recursivamente.
