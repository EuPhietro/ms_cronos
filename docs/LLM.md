# LLM Repository Guide

Este documento existe para acelerar a leitura do repositorio por LLMs, agentes
e assistentes de codigo.

## Resumo Curto

- linguagem principal: Python;
- dominio atual: SharePoint + Microsoft Graph;
- camada principal implementada: `core/`;
- objetivo atual: abstrair o SDK depois de validar navegacao e uploads pequeno
  e grande;
- objetivo futuro: API reutilizavel e upload de diretorios locais complexos.

## O Que Ja Existe

O Core ja implementa:

- autenticacao Graph;
- resolucao de site por URL humana;
- listagem de drives;
- obtencao do drive padrao;
- obtencao da raiz do drive;
- listagem de filhos imediatos;
- busca de filhos por nome, id e `web_url`;
- busca tipada de pasta e arquivo;
- criacao de pasta remota;
- garantia de caminho remoto;
- upload pequeno com conflito `fail`, `rename` ou `replace`;
- upload grande com upload session e chunks;
- parse de models do SDK para models internos;
- traducao inicial de `ODataError`.

Ainda faltam:

- wrappers para os novos tipos do SDK;
- factory/parser para o resultado do upload grande;
- erros semanticos do novo fluxo;
- leitura recursiva de arvore local;
- upload completo de diretorios.

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

Hierarquia remota relevante:

```text
Tenant
  -> Site
  -> Drive
  -> DriveItem
     -> Folder
     -> File
```

Mapeamento para contratos internos:

```text
SharePoint URL
  -> SiteRef
  -> DriveRef
  -> DriveItemRef
```

## Filosofia do Core

Ao mexer no projeto, preserve estas regras:

1. expor contratos pequenos e semanticos;
2. nao vazar envelopes crus do SDK para fora do Core;
3. centralizar parse e traducao de erros em `core.parse`;
4. deixar `core.sharepoint` como orquestrador, nao como model paralelo do SDK.

## Modulos e Responsabilidades

### `core/models.py`

Contem:

- `GraphCredentials`
- `SiteRef`
- `DriveRef`
- `DriveItemRef`
- `LocalFile`
- `StagingContentUpload`
- `UploadFileResult`
- colecoes genericas e colecoes concretas

Use este modulo quando a tarefa envolver:

- novo model interno;
- colecao semantica;
- contrato de retorno do Core.

### `core/errors.py`

Contem a hierarquia semantica de erros do projeto.

Use este modulo quando a tarefa envolver:

- novo erro de dominio;
- refinamento de mensagens de erro;
- mapeamento de falhas do Graph para erros internos.

### `core/graph_client.py`

Contem o manager do Graph.

Use este modulo quando a tarefa envolver:

- credenciais;
- scopes;
- criacao ou fechamento do `GraphServiceClient`.

### `core/parse.py`

Contem adaptadores entre SDK e Core.

Use este modulo quando a tarefa envolver:

- parse de `Site`, `Drive` ou `DriveItem`;
- parse de collection responses;
- parse de `ODataError`;
- parse de caminho local para `LocalFile`.

### `core/urls.py`

Contem helpers de URL e rotas Graph.

Use este modulo quando a tarefa envolver:

- validacao de URL do SharePoint;
- montagem de URL de resolucao;
- montagem de fragmento ou URL de upload pequeno.

### `core/builders.py`

Contem builders de payload e staging.

Use este modulo quando a tarefa envolver:

- body para criacao de pasta;
- staging de upload pequeno;
- validacao de `conflict_behavior`;
- validacao de nome remoto.

### `core/sharepoint.py`

Contem o servico principal.

Use este modulo quando a tarefa envolver:

- navegacao remota;
- criacao de pasta;
- upload pequeno;
- fluxo de negocio sobre site, drive e item.

### `core/utils.py`

Contem helpers pequenos e reutilizaveis.

Hoje o destaque e:

- `rename_with_uuid(...)`

## API Atual Mais Importante

O `SharePointService` hoje expoe principalmente:

- `resolve_site(...)`
- `list_site_drives(...)`
- `get_default_drive(...)`
- `get_drive(...)`
- `get_drive_root(...)`
- `list_children(...)`
- `find_child_by_name(...)`
- `find_child_by_id(...)`
- `find_child_by_web_url(...)`
- `find_child_folder(...)`
- `find_child_file(...)`
- `find_child_folder_by_id(...)`
- `find_child_folder_web_url(...)`
- `create_folder(...)`
- `ensure_remote_folder_path(...)`
- `upload_small_file(...)`
- `upload_large_file(...)`

## Exemplo Minimo

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

        target = await sharepoint.ensure_remote_folder_path(
            drive,
            root,
            ("datasets", "2026", "06"),
        )

        local_file = LocalFile.from_uri("/tmp/relatorio.csv")
        result = await sharepoint.upload_small_file(
            drive,
            target,
            local_file,
            conflict_behavior="rename",
        )

        print(result.remote_name)


asyncio.run(main())
```

## Regras Operacionais Importantes

### Navegacao

- `list_children(...)` nao e recursivo;
- buscas operam sobre filhos imediatos;
- um `drive_item_id` identifica o item dentro do drive informado.

### Upload pequeno

- limite atual: `250_000_000` bytes;
- `fail` levanta erro se o nome remoto ja existir;
- `rename` gera novo nome preservando extensao;
- `replace` substitui o conteudo do arquivo existente.

### Upload grande

- cria uma upload session por path remoto;
- envia o stream por `LargeFileUploadTask`;
- usa `MAX_CHUNK_SIZE` como limite de cada trecho;
- ainda retorna `UploadResult[DriveItem]` do SDK.

## Direcao da Fase 2

A API publica deve aceitar e devolver apenas tipos do MS Cronos.

```text
public API -> internal wrappers/adapters -> Graph SDK
```

Tipos do SDK atualmente visiveis no fluxo grande:

- `DriveItem`
- `DriveItemUploadableProperties`
- `UploadSession`
- `UploadResult`
- `CreateUploadSessionPostRequestBody`
- `LargeFileUploadTask`

Esses tipos devem ficar restritos aos builders, parsers e wrappers internos.

### Erros

O Core prefere levantar erros internos, por exemplo:

- `SiteResolutionError`
- `DriveNotFoundError`
- `DriveItemNotFoundError`
- `NotAFolderError`
- `NotAFileError`
- `FileAlreadyExistError`
- `FileVeryLargeError`
- `SmallFileUploadError`

## Heuristicas Para Novas Tarefas

Se a tarefa mencionar:

- "novo model" -> `core/models.py`
- "novo parse do SDK" -> `core/parse.py`
- "nova URL Graph" -> `core/urls.py`
- "novo body de envio" -> `core/builders.py`
- "nova operacao SharePoint" -> `core/sharepoint.py`
- "novo erro" -> `core/errors.py`

## Arquivos Fonte de Verdade

Para contexto humano:

- `README.md`
- `docs/SPEC.md`

Para contexto tecnico atual:

- `core/models.py`
- `core/sharepoint.py`
- `core/parse.py`
