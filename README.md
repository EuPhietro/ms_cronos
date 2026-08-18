# MS Cronos

MS Cronos e um wrapper Python assincrono para navegacao e upload de arquivos em
bibliotecas de documentos do SharePoint por meio do Microsoft Graph.

O projeto oferece modelos de dominio pequenos e uma API de alto nivel para que
aplicacoes consumidoras nao precisem lidar diretamente com `DriveItem`,
envelopes OData, URLs de continuacao ou detalhes de upload do SDK.

> Status: **Beta 1 (`0.1.0b1`)**. Navegacao remota, scanner local, staging,
> upload sequencial de arvores, checkpoint persistente e progresso observavel
> estao funcionais. Mudancas incompativeis ainda podem ocorrer antes da `1.0`.

## Recursos

- autenticacao de aplicacao com Microsoft Entra ID;
- resolucao de um site pela URL humana do SharePoint;
- listagem paginada de bibliotecas de documentos;
- obtencao da biblioteca padrao e de sua raiz;
- iteracao paginada e acumulacao de itens remotos;
- filtros semanticos para arquivos e pastas;
- busca de itens por nome, ID e URL web;
- criacao e garantia de caminhos remotos;
- upload direto de arquivos pequenos;
- upload de arquivos grandes por sessao e chunks;
- politicas de conflito `fail`, `rename` e `replace`;
- scanner recursivo de diretorios locais baseado em `Path.walk()`;
- snapshot plano com contadores de arquivos, diretorios, niveis e bytes;
- preservacao de diretorios vazios no snapshot local;
- upload top-down de arvores completas com cache dos filhos remotos;
- checkpoint JSON atomico e resultado parcial para retomada de arvores;
- callback de progresso e cancelamento cooperativo;
- retry com `Retry-After` e backoff para falhas transitorias do Graph;
- validacao centralizada de nomes e fragmentos remotos;
- traducao de erros OData para excecoes do Core.

## Modelo Conceitual

O SDK do Microsoft Graph trabalha com `Site`, `Drive` e `DriveItem`. A API de
alto nivel do MS Cronos usa nomes orientados ao dominio:

```text
Microsoft Graph             MS Cronos
---------------             ---------
Site                  ->     SharePointSite
Drive                 ->     DocumentLibrary
DriveItem             ->     SharePointItem
arquivo no disco      ->     LocalFile
diretorio no disco    ->     DirectoryLevel
arvore local          ->     FilesystemTree
plano de upload       ->     StagingFilesystemTree
resultado de upload   ->     FileUploadResult
resultado da arvore   ->     TreeUploadResult
```

As operacoes de `SharePointService` convertem objetos e envelopes do SDK antes
de devolve-los ao consumidor.

## Requisitos

- Python 3.12 ou superior (`LocalFileSystemScanner` usa `Path.walk()`);
- uma aplicacao registrada no Microsoft Entra ID;
- permissoes de aplicacao adequadas para os sites e arquivos acessados;
- consentimento administrativo quando exigido pelo tenant.

## Instalacao

A partir de um checkout do repositorio:

```bash
python -m venv venv
source venv/bin/activate
python -m pip install --upgrade pip
python -m pip install -e ".[dev]"
```

No PowerShell:

```powershell
python -m venv venv
.\venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
python -m pip install -e ".[dev]"
```

O pacote possui metadados de distribuicao em `pyproject.toml`, mas ainda nao
foi publicado no PyPI.

O codigo-fonte e distribuido publicamente em
[github.com/EuPhietro/ms_cronos](https://github.com/EuPhietro/ms_cronos).

## Configuracao

Use `.env.example` como referencia para criar um arquivo `.env` local:

```env
CLIENT_ID=identificador-da-aplicacao
CLIENT_SECRET=segredo-da-aplicacao
CLIENT_TENANT=identificador-do-tenant
UPLOAD_SOURCE=/dados/documentos
UPLOAD_CHECKPOINT=.ms-cronos-upload.json
SHAREPOINT_SITE_URL=https://tenant.sharepoint.com/sites/Financeiro
SHAREPOINT_LIBRARY=Documents
SHAREPOINT_TARGET_ROOT=backup
```

Nunca envie `.env`, client secrets, tokens ou URLs pre-autenticadas de upload
para o controle de versao.

## Inicio Rapido

```python
import asyncio
from os import getenv

from dotenv import load_dotenv

from core import (
    GraphClientManager,
    GraphCredentials,
    LocalFile,
    SharePointService,
)


async def main() -> None:
    load_dotenv()

    credentials = GraphCredentials(
        client_id=getenv("CLIENT_ID", ""),
        client_secret=getenv("CLIENT_SECRET", ""),
        tenant_id=getenv("CLIENT_TENANT", ""),
    )

    async with GraphClientManager(credentials) as manager:
        sharepoint = SharePointService(manager)

        site = await sharepoint.resolve_site(
            "https://tenant.sharepoint.com/sites/Financeiro"
        )
        library = await sharepoint.get_default_drive(site)
        root = await sharepoint.get_drive_root(library)

        destination = await sharepoint.ensure_remote_folder_path(
            library,
            root,
            ("relatorios", "2026"),
        )

        result = await sharepoint.upload(
            library,
            destination,
            LocalFile.from_uri("/tmp/relatorio.csv"),
            conflict_behavior="rename",
        )

        print(result.item.id)
        print(result.remote_name)


if __name__ == "__main__":
    asyncio.run(main())
```

`SharePointService.upload` escolhe automaticamente o fluxo direto ou o fluxo
por sessao conforme o tamanho do arquivo.

## Scanner Local

`LocalFileSystemScanner` percorre uma raiz sem ler o conteudo binario dos
arquivos. O resultado e um snapshot plano: cada diretorio encontrado vira um
`DirectoryLevel` com seus arquivos e subdiretorios imediatos.

```python
from core import LocalFileSystemScanner

scanner = LocalFileSystemScanner()
tree = scanner.scan(
    "/tmp/documentos",
    allow_empty="deny",
    sort_entries=True,
)

print(tree.root.path)
print(tree.total_files)
print(tree.total_size)
print(tree.total_levels)
print(tree.total_subdirectories)
```

`allow_empty="deny"` rejeita arquivos de zero bytes. Use `"allow"` quando o
snapshot precisar representar arquivos vazios. `sort_entries=True` torna a
ordem deterministica; sem essa opcao, a ordem segue o sistema de arquivos.

O scanner nao conhece sites, bibliotecas, caminhos remotos nem politicas de
conflito. `StagingTreeBuilder` faz essa associacao sem executar chamadas de
rede.

## Upload De Arvores

```python
import asyncio
from pathlib import PurePosixPath

cancel_event = asyncio.Event()


def report_progress(progress: TreeUploadProgress) -> None:
    print(progress.phase, progress.completed_files, progress.total_files)


staging_tree = StagingTreeBuilder().build_staging_tree(
    tree,
    conflict_behavior="replace",
    target_root=PurePosixPath("backup/documentos"),
)

result = await sharepoint.upload_tree(
    root,
    library,
    staging_tree,
    checkpoint_path=".ms-cronos-upload.json",
    checkpoint_interval=100,
    progress_callback=report_progress,
    cancel_event=cancel_event,
)

print(result.total_uploaded_files)
```

O executor cria as pastas de cima para baixo, lista cada nivel remoto uma vez
por execucao e reutiliza esse indice durante os uploads. Quando
`checkpoint_path` existe, ele e carregado automaticamente. O checkpoint e
vinculado ao caminho local, biblioteca, pai remoto e `target_root`, impedindo
retomada acidental em outro destino. Uma impressao digital de caminhos, nomes,
tamanhos e politicas de conflito tambem rejeita uma arvore de staging
estruturalmente diferente. `TreeUploadError.partial_result` continua disponivel
para tratamento em memoria.

## Navegacao Remota

### Resolver um site

```python
site = await sharepoint.resolve_site("https://tenant.sharepoint.com/sites/Financeiro")

print(site.id)
print(site.display_name)
print(site.web_url)
```

### Listar bibliotecas

Por padrao, `list_site_drives` percorre todas as paginas:

```python
libraries = await sharepoint.list_site_drives(site)

for library in libraries:
    print(library.id, library.name, library.drive_type)
```

Tambem e possivel limitar o numero de paginas processadas:

```python
libraries = await sharepoint.list_site_drives(
    site,
    pagination=True,
    max_pages=2,
)
```

### Acumular todos os filhos

`list_children` consome a paginacao e retorna uma unica colecao:

```python
root = await sharepoint.get_drive_root(library)
children = await sharepoint.list_children(library, root)

for child in children:
    print(child.name, child.is_folder, child.is_file)
```

Um predicado opcional pode filtrar os itens durante a acumulacao:

```python
folders = await sharepoint.list_children(
    library,
    root,
    filter=lambda item: item.is_folder,
)
```

### Consumir pagina por pagina

Para evitar acumular todos os itens em memoria, use `iter_children`:

```python
async for page in sharepoint.iter_children(library, root):
    print(f"Itens nesta pagina: {page.counter}")

    for item in page:
        print(item.name)
```

O `PageIterator` do SDK permanece interno; o consumidor recebe somente
`SharePointItemCollection`.

### Buscar arquivos e pastas

```python
folder = await sharepoint.find_folder_by_name(
    library,
    root,
    "Relatorios",
)

file = await sharepoint.find_file_by_name(
    library,
    root,
    "balanco.pdf",
)
```

As buscas retornam `None` quando nao encontram um item correspondente.

## Upload

```python
local_file = LocalFile.from_uri("/tmp/backup.zip")

result = await sharepoint.upload(
    library,
    root,
    local_file,
    conflict_behavior="replace",
)

print(result.source_path)
print(result.remote_name)
print(result.item.web_url)
```

Politicas de conflito:

- `fail`: interrompe se o nome remoto ja existir;
- `rename`: cria o recurso com um novo nome;
- `replace`: substitui o conteudo do arquivo remoto existente.

O retorno de ambos os fluxos e `FileUploadResult`; detalhes como
`UploadSession`, `LargeFileUploadTask` e `DriveItem` nao fazem parte do
resultado publico.

## Colecoes Semanticas

As listagens retornam colecoes tipadas em vez de listas cruas:

- `SharePointSiteCollection`;
- `DocumentLibraryCollection`;
- `SharePointItemCollection`;
- `LocalFileCollection`;
- `LocalFolderCollection`;
- `DirectoryLevelCollection`.

Elas suportam iteracao, `len`, acesso por indice, slices e operacoes auxiliares:

```python
print(children.counter)
print(children.is_empty)
print(children.first())
print(children.to_list())
```

## Acesso Avancado Ao SDK

`GraphClientManager.client` e uma extensao intencional para operacoes ainda nao
cobertas por `SharePointService`:

```python
async with GraphClientManager(credentials) as manager:
    raw_response = await manager.client.sites.get()
```

Ao usar essa propriedade, o codigo consumidor passa a depender diretamente das
assinaturas e dos modelos do Microsoft Graph SDK. Para uso comum, prefira
`SharePointService`.

## Tratamento De Erros

As excecoes de dominio derivam de uma raiz comum. Entre as especializacoes
disponiveis estao:

- `SiteResolutionError`;
- `DriveNotFoundError`;
- `DriveItemNotFoundError`;
- `GraphPermissionError`;
- `GraphAuthenticationError`;
- `GraphResourceConflictError`;
- `LocalPathNotFoundError`;
- `InvalidConflictBehaviorError`;
- `UploadError`.

Exemplo:

```python
from core import UploadError

try:
    result = await sharepoint.upload(library, root, local_file)
except UploadError as error:
    print(f"Operacao nao concluida: {error}")
```

## Arquitetura

```text
Aplicacao consumidora
  -> models, colecoes, erros e SharePointService
  -> parsers, builders e helpers internos
  -> Microsoft Graph SDK
  -> SharePoint
```

Modulos principais:

```text
core/
├── models.py        # contratos e colecoes semanticas
├── errors.py        # hierarquia de erros do dominio
├── filesystem.py    # scanner e snapshot da arvore local
├── graph_client.py  # autenticacao e ciclo de vida do client
├── sharepoint.py    # operacoes de alto nivel
├── parse.py         # adaptacao entre SDK e Core
├── builders.py      # montagem de payloads internos
├── urls.py          # validacao e construcao de rotas
└── utils.py         # utilitarios pequenos
```

Os parsers e builders existem para uso interno. Eles nao devem ser considerados
parte da API estavel do pacote.

## Desenvolvimento

`main.py` funciona como verificacao manual do fluxo completo:

```bash
python main.py
```

Execute a suite reproduzivel com:

```bash
python -m unittest discover -v
ruff check core tests main.py
pyright --pythonpath venv/bin/python core
```

Ela cobre scanner, staging, paginação, conflitos, retry, validacao remota,
cache, checkpoint, progresso, cancelamento e API publica. O teste Graph somente
leitura fica desativado por padrao. Para executa-lo em um tenant controlado:

```bash
MS_CRONOS_RUN_INTEGRATION=1 python -m unittest \
  tests.test_integration_readonly -v
```

O scanner tambem foi exercitado manualmente com uma arvore de 45.779 arquivos,
21.862 niveis e 554.889.293 bytes.

## Roadmap

- adicionar concorrencia limitada com controle adaptativo de throttling;
- ampliar a matriz de integracao com uploads em um destino descartavel;
- endurecer a retomada de uploads grandes interrompidos no meio de um chunk;
- modelar uma arvore remota navegavel;
- preparar publicacao da distribuicao.

## Seguranca

- use o menor conjunto de permissoes compativel com a operacao;
- nao registre secrets, tokens ou URLs de upload;
- nao reutilize URLs de sessao fora do fluxo que as criou;
- trate os IDs e metadados do tenant como informacao sensivel;
- revise logs antes de compartilha-los publicamente.

## Documentacao Complementar

- [Especificacao tecnica](docs/SPEC.md)
- [Guia de contexto para LLMs](docs/LLM.md)

## Licenca

Distribuido sob a licenca MIT. Consulte [`LICENSE`](LICENSE).
