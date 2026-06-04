# LLM Repository Guide

Este documento existe para acelerar a leitura automatica do repositorio por
LLMs, agentes e assistentes de codigo.

## Projeto

Nome: `MS Cronos`

Objetivo principal:
automatizar navegacao e upload de conteudo local para bibliotecas de documentos
do SharePoint via Microsoft Graph.

Objetivo de medio prazo do Core:
permitir upload de diretorios locais complexos preservando a estrutura remota,
sem expor a verbosidade do SDK oficial para o restante do sistema.

## Estado Atual

O repositorio esta concentrado na camada `core/`.

Ja existe implementacao funcional para:

- autenticacao com Microsoft Graph;
- resolucao de site SharePoint por URL humana;
- listagem de drives de um site;
- obtencao do drive padrao;
- obtencao da raiz do drive;
- listagem de filhos de uma pasta remota;
- busca imediata de filhos por nome, id e `web_url`;
- criacao de pastas remotas;
- garantia de caminho remoto com criacao incremental;
- upload de arquivos pequenos;
- parse de models do SDK para models internos;
- traducao inicial de `ODataError` para erros semanticos do Core.

Ainda faltam, em especial:

- upload de arquivos grandes com upload session;
- leitura e mapeamento de arvores locais;
- upload completo de diretorios locais complexos;
- consolidacao final dos fluxos de navegacao e sincronizacao.

## Estrutura Relevante

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

## Ordem Recomendada de Leitura

Para entender o repositorio rapidamente:

1. `README.md`
2. `docs/SPEC.md`
3. `core/__init__.py`
4. `core/models.py`
5. `core/errors.py`
6. `core/parse.py`
7. `core/urls.py`
8. `core/builders.py`
9. `core/sharepoint.py`
10. `main.py`

## Mapa Conceitual

Hierarquia remota importante no SharePoint/Graph:

```text
Tenant
  -> Site
  -> Drive
  -> DriveItem
     -> Folder
     -> File
```

Leitura equivalente dentro do Core:

```text
sharepoint human url
  -> SiteRef
  -> DriveRef
  -> DriveItemRef
```

## Filosofia do Core

O Core tenta manter quatro regras:

1. expor contratos pequenos e semanticos;
2. nao vazar envelopes crus do SDK para fora do Core;
3. centralizar parse e traducao de erros;
4. deixar `sharepoint.py` como servico de orquestracao, nao como deposito de
   estruturas cruas do Graph.

## Modulos do Core

### `core/models.py`

Define os contratos internos do projeto.

Itens mais importantes:

- `GraphCredentials`
- `SiteRef`
- `DriveRef`
- `DriveItemRef`
- `LocalFile`
- `StagingContentUpload`
- `UploadResult`

Tambem define a familia de colecoes:

- `Collection_[T]`
- `FrozenCollection[T]`
- `MutableCollection[T]`
- `SiteRefCollection`
- `DriveRefCollection`
- `DriveItemCollection`
- `LocalFileCollection`
- `StagingUpdateContentCollection`

Observacao:
as colecoes sao parte importante do design semantico do projeto. O repositorio
evita retornar listas cruas quando existe uma colecao concreta mais expressiva.

### `core/errors.py`

Contem a hierarquia semantica de erros do projeto.

Intencao:
quem consome o Core deve lidar com erros do dominio interno, e nao com
excecoes cruas do SDK do Microsoft Graph.

Exemplos importantes:

- `SiteResolutionError`
- `DriveNotFoundError`
- `DriveItemNotFoundError`
- `FolderNotFoundError`
- `NotAFolderError`
- `NotAFileError`
- `FileAlreadyExistError`
- `FileVeryLargeError`
- `SmallFileUploadError`
- `LargeFileUploadError`

### `core/graph_client.py`

Responsavel por:

- validar credenciais;
- criar o `GraphServiceClient`;
- encapsular o ciclo de vida do `ClientSecretCredential`.

### `core/parse.py`

Responsavel por adaptar dados externos para contratos internos.

Principais grupos de funcoes:

- parse de `Site`, `Drive` e `DriveItem`;
- parse de collection responses do SDK;
- parse de caminho local para `LocalFile`;
- parse de `ODataError` para erros internos.

Regra importante:
se um dado veio do SDK e precisa sair do Core, o caminho preferido e passar por
`parse.py`.

### `core/urls.py`

Responsavel por:

- validar URL humana do SharePoint;
- montar a URL especial de resolucao de site para o Graph;
- montar fragmentos e URLs de upload de conteudo.

### `core/builders.py`

Responsavel por construir payloads e staging objects usados pelo servico.

Pontos importantes:

- `build_folder_drive_item(...)`
- `build_upload_content(...)`

O builder de upload pequeno prepara:

- `target_path`
- `conflict_behavior`
- validacao de nome remoto

### `core/sharepoint.py`

Servico principal do Core.

Ele orquestra chamadas ao Graph e retorna apenas modelos internos.

Metodos importantes ja implementados:

- `resolve_site`
- `list_site_drives`
- `get_default_drive`
- `get_drive`
- `get_drive_root`
- `list_children`
- `find_child_by_name`
- `find_child_by_id`
- `find_child_by_web_url`
- `find_child_folder`
- `find_child_file`
- `create_folder`
- `ensure_remote_folder_path`
- `upload_small_file`

### `core/utils.py`

Helpers pequenos e reutilizaveis.

Hoje o principal exemplo e:

- `rename_with_uuid(...)`

Esse helper preserva extensao ao gerar um nome alternativo.

## API Publica

Quando outro modulo do projeto quiser consumir o Core, a preferencia e importar
de `core.__init__`, nao de submodulos internos, salvo quando estiver
implementando o proprio Core.

Exemplo:

```python
from core import GraphClientManager, GraphCredentials, SharePointService
```

## Fluxo Feliz Atual

O fluxo funcional atual e aproximadamente:

```text
GraphCredentials
  -> GraphClientManager
  -> SharePointService
  -> resolve_site(url)
  -> get_default_drive(site)
  -> get_drive_root(drive)
  -> list_children(drive, root)
  -> ensure_remote_folder_path(...)
  -> upload_small_file(...)
```

## Regras e Convencoes Importantes

### 1. Parse centralizado

Evite adaptar manualmente `Site`, `Drive` e `DriveItem` dentro de
`sharepoint.py` quando um parser do modulo `parse.py` puder ser usado.

### 2. Erros semanticos

Se uma operacao falha por regra de negocio, o retorno esperado e um erro do
Core, nao uma excecao generica.

### 3. IDs do Graph

Para navegar entre niveis remotos, geralmente nao e necessario concatenar ids
ancestrais na URL. Um `drive_id` identifica o drive e um `drive_item_id`
identifica o item dentro daquele drive.

### 4. Small upload

O fluxo atual de upload pequeno trata conflitos via `conflict_behavior`.

Valores aceitos:

- `fail`
- `rename`
- `replace`

### 5. Limite atual para arquivo pequeno

O Core hoje considera arquivo grande acima de `250_000_000` bytes para o fluxo
de upload pequeno.

### 6. `main.py` nao e API

`main.py` funciona como laboratorio manual e roteiro de validacao local.
Ele nao representa a interface publica final do projeto.

## Ponto de Entrada Mental para Novas Implementacoes

Se a tarefa for:

- modelagem de dados: comece por `core/models.py`;
- traducao de SDK para Core: comece por `core/parse.py`;
- montagem de payload ou staging: comece por `core/builders.py`;
- navegacao, criacao ou upload remoto: comece por `core/sharepoint.py`;
- construcao de URLs do Graph: comece por `core/urls.py`;
- novo erro de dominio: comece por `core/errors.py`.

## Lacunas Conhecidas

As areas que ainda devem evoluir com mais cuidado sao:

- upload grande em partes;
- politica completa de conflitos para todas as operacoes;
- leitura de diretorios locais e representacao de arvore;
- consolidacao do fluxo de upload de diretorios inteiros;
- possiveis refinamentos de comentarios, mensagens e assinaturas.

## Exemplo de Leitura Rapida por LLM

Se um agente precisar responder rapidamente "onde implementar X?", use estas
heuristicas:

- "novo model interno" -> `core/models.py`
- "novo parse do SDK" -> `core/parse.py`
- "novo erro semantico" -> `core/errors.py`
- "nova URL ou fragmento Graph" -> `core/urls.py`
- "novo payload/body de envio" -> `core/builders.py`
- "nova operacao SharePoint" -> `core/sharepoint.py`
- "exemplo manual de uso" -> `main.py`

## Arquivos Fonte de Verdade

Para contexto humano e de produto:

- `README.md`
- `docs/SPEC.md`

Para contexto tecnico atual:

- `core/models.py`
- `core/sharepoint.py`
- `core/parse.py`
