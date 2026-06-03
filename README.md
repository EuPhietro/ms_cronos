# MS Cronos

MS Cronos e um projeto Python para automatizar navegacao e upload de conteudo
local para bibliotecas de documentos do SharePoint via Microsoft Graph.

O foco atual esta no Core: uma camada pequena, tipada e semantica que esconde a
verbosidade do SDK oficial e expoe operacoes previsiveis de site, drive, pasta
e arquivo.

## Objetivo

O objetivo final do projeto e permitir upload de diretorios locais complexos
para o SharePoint, preservando a estrutura de pastas e reduzindo o contato do
restante do sistema com detalhes do Microsoft Graph.

Em termos práticos, o Core deve permitir fluxos como:

```text
diretorio local
  -> resolver site
  -> escolher drive
  -> obter root
  -> encontrar ou criar pastas remotas
  -> enviar arquivos pequenos e grandes
```

## Estado Atual

O projeto ja possui:

- autenticacao com `ClientSecretCredential`;
- construcao de `GraphServiceClient`;
- resolucao de site SharePoint por URL humana;
- listagem de drives de um site;
- obtencao do drive padrao;
- obtencao de drive especifico por id;
- resolucao da raiz do drive;
- listagem dos filhos de um `DriveItem` pasta;
- busca imediata de filhos por nome, id e URL web;
- parse semantico entre models do SDK e models internos;
- traducao inicial de `ODataError` para erros do Core.

## Estrutura Atual

```text
ms_cronos/
├─ core/
│  ├─ __init__.py
│  ├─ errors.py
│  ├─ graph_client.py
│  ├─ models.py
│  ├─ parse.py
│  ├─ sharepoint.py
│  └─ urls.py
├─ docs/
│  └─ SPEC.md
├─ main.py
├─ README.md
└─ requirements.txt
```

## Mapa Conceitual

A hierarquia principal do SharePoint para este projeto e:

```text
Tenant
  -> Site
  -> Drive (document library)
  -> DriveItem
     -> Folder
     -> File
```

No Core isso vira:

```text
URL humana
  -> SiteRef
  -> DriveRef
  -> DriveItemRef
```

## Modulos do Core

### `core/models.py`

Contem:

- modelos semanticos (`SiteRef`, `DriveRef`, `DriveItemRef`, `LocalFile`, `UploadResult`);
- colecoes genericas (`Collection_`, `FrozenCollection`, `MutableCollection`);
- colecoes concretas (`SiteRefCollection`, `DriveRefCollection`, `DriveItemCollection`, `LocalFileCollection`).

### `core/graph_client.py`

Responsavel por:

- validar credenciais;
- criar o `GraphServiceClient`;
- encapsular ciclo de vida do `ClientSecretCredential`.

### `core/parse.py`

Responsavel por:

- converter `Site`, `Drive` e `DriveItem` do SDK em modelos internos;
- converter envelopes `...CollectionResponse` em colecoes concretas do Core;
- adaptar `ODataError` para a hierarquia de erros interna;
- adaptar caminhos locais para `LocalFile`.

### `core/sharepoint.py`

Contem o servico principal, atualmente com operacoes de leitura e navegacao:

- `resolve_site`
- `list_site_drives`
- `get_default_drive`
- `get_drive`
- `get_drive_root`
- `list_children`
- `find_child_by_name`
- `find_child_by_id`
- `find_child_by_web_url`

### `core/urls.py`

Contem utilitarios pequenos para:

- validar URLs do SharePoint;
- construir rotas de resolucao do Graph.

### `core/errors.py`

Contem a hierarquia semantica de erros do projeto, evitando vazar excecoes do
SDK para as camadas superiores.

## Fluxo Ja Validado

O `main.py` ainda funciona como laboratorio manual. O fluxo ja exercitado e:

1. montar `GraphCredentials`;
2. criar `GraphClientManager`;
3. instanciar `SharePointService`;
4. resolver um site SharePoint;
5. listar drives do site;
6. obter drive padrao ou drive especifico;
7. obter o root do drive;
8. listar filhos da raiz;
9. buscar filhos imediatos por nome, id ou URL web.

## Proximas Implementacoes

Para chegar ao objetivo de upload de diretorios locais complexos, as proximas
funcoes do Core sao:

1. `find_child`
2. `find_child_folder`
3. `create_folder`
4. `upload_file`
5. `upload_large_file`
6. uma camada de leitura/orquestracao para arvore local
7. um fluxo de alto nivel de upload de diretorio

## API Publica Atual do Servico

Hoje o desenho do servico esta mais proximo de:

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
```

## Microsoft Graph

Rotas centrais do projeto:

```http
GET /sites/{hostname}:/{site-path}
GET /sites/{site-id}/drives
GET /sites/{site-id}/drive
GET /drives/{drive-id}
GET /drives/{drive-id}/root
GET /drives/{drive-id}/items/{item-id}/children
POST /drives/{drive-id}/items/{item-id}/children
PUT /drives/{drive-id}/items/{item-id}:/{filename}:/content
```

## Autenticacao

O projeto usa autenticacao de aplicacao com `ClientSecretCredential`.

Variaveis esperadas no `.env`:

```env
CLIENT_ID=
CLIENT_SECRET=
CLIENT_TENANT=
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

## Uso Atual

```bash
python main.py
```

O `main.py` ainda e um arquivo de validacao manual e aprendizado do fluxo real
contra o Graph.

## Principios do Projeto

- o Core deve ser pequeno e testavel;
- o SharePoint deve ser exposto por contratos semanticos, nao por envelopes do SDK;
- o parser deve concentrar adaptacao de models e erros;
- a camada de servico deve orquestrar chamadas, nao inventar estruturas paralelas;
- a evolucao futura para daemon e jobs deve reutilizar o Core sem acoplamento.

## Leitura Complementar

As especificacoes mais detalhadas do projeto estao em [docs/SPEC.md](docs/SPEC.md).
