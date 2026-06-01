# MS Cronos Specification

Este documento descreve a especificacao geral do MS Cronos e o estado atual do
seu Core.

## Visao Geral

MS Cronos deve automatizar navegacao e upload de conteudo local para o
SharePoint usando Microsoft Graph.

O objetivo de produto de mais alto nivel e permitir upload de diretorios locais
complexos, preservando a estrutura remota com o menor acoplamento possivel ao
SDK oficial.

Fluxo alvo:

```text
diretorio local
  -> leitura de arvore local
  -> resolve site
  -> escolhe drive
  -> encontra ou cria pastas remotas
  -> envia arquivos
  -> retorna resultado estavel
```

## Objetivos

- Criar um wrapper Python pequeno para SharePoint via Microsoft Graph.
- Expor contratos semanticos e tipados para site, drive, pasta e arquivo.
- Esconder a verbosidade do SDK oficial.
- Preparar o terreno para upload de diretorios locais complexos.
- Manter o Core reutilizavel por futuras camadas de daemon e jobs.

## Nao Objetivos

Na fase atual, o projeto ainda nao busca:

- sincronizacao bidirecional;
- UI;
- monitoramento automatico de diretorios;
- suporte generico a todos os recursos do Graph;
- camada HTTP final;
- daemon Windows finalizado.

## Camadas do Projeto

Arquitetura desejada:

```text
Presentation / API
  -> recebe comandos externos e cria jobs

Application / Jobs
  -> define contratos de execucao e estados

Daemon / Worker
  -> executa jobs em background

Core
  -> implementa operacoes de SharePoint/Drive

Infrastructure
  -> configuracao, credenciais, logging e persistencia
```

A fase atual esta concentrada no **Core**.

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

## Core

O Core e a camada principal da primeira fase.

### Responsabilidades

- criar cliente autenticado do Microsoft Graph;
- resolver site por URL humana do SharePoint;
- listar drives de um site;
- obter drive padrao;
- obter drive especifico;
- obter root do drive;
- listar filhos de um item pasta;
- encontrar arquivo ou pasta por nome;
- criar pasta remota;
- enviar arquivos pequenos e grandes;
- traduzir modelos e erros do SDK para contratos internos.

### Fora do escopo do Core

- ler `.env` como responsabilidade de negocio;
- expor API HTTP;
- orquestrar loop de daemon;
- armazenar jobs;
- vazar `ODataError` ou envelopes crus do SDK.

## Modelos do Core

O projeto trabalha com contratos internos leves:

### Itens

- `GraphCredentials`
- `SiteRef`
- `DriveRef`
- `DriveItemRef`
- `LocalFile`
- `UploadResult`

### Colecoes

- `Collection_[T]`
- `FrozenCollection[T]`
- `MutableCollection[T]`
- `SiteRefCollection`
- `DriveRefCollection`
- `DriveItemCollection`
- `LocalFileCollection`

## Modulos do Core

### `core/graph_client.py`

Responsavel por:

- validar credenciais;
- criar `GraphServiceClient`;
- encapsular o `ClientSecretCredential`;
- oferecer um manager reutilizavel para o servico.

### `core/models.py`

Responsavel por:

- contratos semanticos do Core;
- colecoes reutilizaveis;
- separacao entre leitura, mutabilidade e imutabilidade.

### `core/parse.py`

Responsavel por:

- converter `Site`, `Drive` e `DriveItem` do SDK em models internos;
- converter envelopes `SiteCollectionResponse`, `DriveCollectionResponse` e
  `DriveItemCollectionResponse` em colecoes concretas;
- converter `Path` local em `LocalFile`;
- traduzir `ODataError` para a hierarquia de erros do Core.

### `core/sharepoint.py`

Responsavel por:

- orquestrar chamadas ao Graph;
- validar envelopes e regras de dominio;
- delegar adaptacao ao modulo `parse`;
- devolver apenas contratos internos.

### `core/urls.py`

Responsavel por:

- validar URLs humanas do SharePoint;
- construir a rota de resolucao usada pelo Graph.

### `core/errors.py`

Responsavel por:

- hierarquia semantica de erros do projeto;
- isolamento de erros do SDK;
- melhor ergonomia de debug para camadas superiores.

## API Publica Atual do Servico

O servico atual se aproxima deste contrato:

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
```

### Operacoes ainda faltantes

As proximas operacoes previstas para o servico sao:

```python
async def find_child(...) -> DriveItemRef | None: ...
async def find_child_folder(...) -> DriveItemRef | None: ...
async def create_folder(...) -> DriveItemRef: ...
async def upload_file(...) -> UploadResult: ...
async def upload_large_file(...) -> UploadResult: ...
```

Depois disso, uma camada de orquestracao maior deve permitir:

```python
async def upload_directory(...) -> ...: ...
```

## Hierarquia Conceitual do SharePoint

Para este projeto, a estrutura mais importante e:

```text
Tenant
  -> Site
  -> Drive (document library)
  -> DriveItem
     -> Folder
     -> File
```

Leitura disso no Core:

```text
URL humana
  -> SiteRef
  -> DriveRef
  -> DriveItemRef
```

## Rotas Principais do Graph

Rotas mais importantes para o estado atual e proximo do projeto:

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

## Fluxo Ja Validado

O `main.py` hoje e um laboratorio manual. O fluxo ja exercitado no projeto e:

1. montar `GraphCredentials`;
2. criar `GraphClientManager`;
3. instanciar `SharePointService`;
4. resolver site por URL humana;
5. listar drives do site;
6. obter drive padrao ou drive especifico;
7. obter o root do drive;
8. listar os filhos do root.

## ODataError e Traducoes

O projeto ja comecou a tratar `ODataError` no parser de erro.

Codigos mais relevantes:

- `accessDenied` -> `GraphPermissionError`
- `authorizationRequestDenied` -> `GraphPermissionError`
- `unauthenticated` -> `GraphAuthenticationError`
- `invalidAuthenticationToken` -> `GraphAuthenticationError`
- `invalidRequest` -> `GraphRequestError`
- `badRequest` -> `GraphRequestError`
- `badArgument` -> `GraphRequestError`
- `nameAlreadyExists` -> `GraphResourceConflictError`
- `conflict` -> `GraphResourceConflictError`
- `itemNotFound` -> erro contextual
- `resourceNotFound` -> erro contextual

## Upload de Diretorios Complexos

Como objetivo final, o projeto deve suportar um fluxo de upload recursivo de
diretorios locais complexos.

Isso implica:

1. leitura de arvore local;
2. preservacao de caminho relativo;
3. busca de pasta remota por nome;
4. criacao de pasta ausente;
5. upload de arquivos pequenos;
6. upload resumivel de arquivos grandes;
7. orquestracao recursiva.

### Primitivas necessarias antes da orquestracao

- `find_child`
- `find_child_folder`
- `create_folder`
- `upload_file`
- `upload_large_file`

## Principios do Projeto

- o Core deve ser pequeno e previsivel;
- parsers devem adaptar, nao orquestrar;
- servicos devem orquestrar, nao vazar SDK;
- erros devem ser semanticos e legiveis;
- a camada atual deve priorizar clareza antes de automacao pesada;
- o objetivo final de diretorios complexos deve surgir a partir de primitivas pequenas.

## Proximos Passos

1. implementar `find_child`;
2. implementar `find_child_folder`;
3. implementar `create_folder`;
4. implementar `upload_file`;
5. implementar `upload_large_file`;
6. modelar leitura recursiva local;
7. orquestrar upload de diretorio.
