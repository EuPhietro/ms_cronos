# MS Cronos - Especificacao Tecnica

## 1. Estado

MS Cronos e um wrapper Python assincrono para Microsoft Graph e SharePoint.
Seu Core fornece tipos semanticos para autenticacao, navegacao remota, criacao
de pastas, upload de arquivos e leitura de arvores locais.

Estado desta especificacao:

- navegacao de sites, bibliotecas e itens: implementada;
- paginacao de bibliotecas e itens: implementada;
- criacao e garantia de caminhos remotos: implementada;
- upload individual pequeno e grande: implementado;
- scanner recursivo do filesystem: implementado;
- staging e upload de diretorios completos: planejados;
- testes automatizados e empacotamento: pendentes.

Esta especificacao descreve o codigo atual. Os modelos do SDK mencionados aqui
sao detalhes internos, exceto quando o consumidor usa deliberadamente
`GraphClientManager.client`.

## 2. Requisitos

- Python 3.12 ou superior, pois o scanner usa `pathlib.Path.walk()`;
- `msgraph-sdk==1.58.0`;
- credenciais de aplicacao do Microsoft Entra ID;
- permissoes de aplicacao compativeis com os recursos acessados;
- consentimento administrativo quando exigido pelo tenant.

## 3. Fronteiras

```text
Aplicacao
  -> modelos, colecoes, erros e servicos do Core
  -> parsers, builders e URLs internas
  -> Microsoft Graph SDK
  -> SharePoint
```

Regras:

1. Metodos publicos de `SharePointService` recebem e retornam modelos do Core.
2. `Site`, `Drive`, `DriveItem`, envelopes OData e `PageIterator` permanecem
   dentro da integracao.
3. `GraphClientManager.client` e uma extensao avancada e, quando usada,
   transfere ao consumidor a dependencia direta do SDK.
4. `LocalFileSystemScanner` nao conhece SharePoint nem executa upload.
5. O futuro staging associa o snapshot local a destinos remotos.

## 4. Hierarquia Remota

```text
SharePointSite
  -> DocumentLibrary
    -> SharePointItem raiz
      -> SharePointItem pasta ou arquivo
```

Um `SharePointItem.id` identifica o item dentro do drive. A navegacao por ID
nao exige concatenar os IDs dos ancestrais. Operacoes por caminho ainda
precisam de um pai ou fragmento remoto coerente.

## 5. Modelos Publicos

### 5.1 Autenticacao

```python
GraphCredentials(
    client_id: str,
    client_secret: str,
    tenant_id: str,
)
```

### 5.2 Recursos remotos

```python
SharePointSite(
    id: str,
    name: str | None = None,
    display_name: str | None = None,
    web_url: str | None = None,
)

DocumentLibrary(
    id: str,
    name: str | None = None,
    web_url: str | None = None,
    drive_type: str | None = None,
)

SharePointItem(
    id: str,
    name: str | None = None,
    web_url: str | None = None,
    is_folder: bool = False,
    is_file: bool = False,
    size: int | None = None,
)
```

### 5.3 Arquivos locais e upload

```python
LocalFile(
    path: Path,
    name: str,
    size: int,
    extension: str,
    allow_empty: Literal["allow", "deny"] = "deny",
)

FileUploadResult(
    item: SharePointItem,
    source_path: Path,
    remote_name: str,
    conflict_behavior: ConflictBehavior,
)
```

`ConflictBehavior` aceita `"fail"`, `"rename"` ou `"replace"`.

Por padrao, `LocalFile` rejeita arquivos vazios. A politica `"allow"` existe
para representar explicitamente arquivos de zero bytes no snapshot.

### 5.4 Snapshot local

```python
RootFolder(path: Path, name: str)
LocalFolder(path: Path, name: str)

DirectoryLevel(
    *,
    path: Path,
    files: LocalFileCollection,
    folders: LocalFolderCollection,
)

FilesystemTree(
    *,
    root: RootFolder,
    levels: DirectoryLevelCollection,
)
```

`FilesystemTree` e plano. Cada `DirectoryLevel` contem somente os arquivos e
subdiretorios imediatos de seu proprio caminho. A relacao hierarquica pode ser
derivada dos paths absolutos e da ordem top-down.

Contadores disponiveis:

- `DirectoryLevel.total_size`;
- `DirectoryLevel.total_files`;
- `DirectoryLevel.total_directories`;
- `FilesystemTree.total_size`;
- `FilesystemTree.total_files`;
- `FilesystemTree.total_levels`;
- `FilesystemTree.total_subdirectories`.

O snapshot guarda metadados, nao os bytes dos arquivos.

## 6. Colecoes

`Collection_[T]` define o contrato comum de leitura:

- `counter`;
- `is_empty`;
- `first()`;
- `to_list()`;
- iteracao;
- `len`;
- acesso por indice e slice;
- teste com `in`.

`FrozenCollection[T]` usa tupla e devolve uma nova instancia em `add`,
`extend`, `remove` e `clear`. `MutableCollection[T]` usa lista e altera a
instancia atual.

Colecoes concretas:

- `SharePointSiteCollection`;
- `DocumentLibraryCollection`;
- `SharePointItemCollection`;
- `LocalFileCollection`;
- `LocalFolderCollection`;
- `DirectoryLevelCollection`;
- `PreparedUploadCollection`.

## 7. Scanner Local

Assinatura:

```python
LocalFileSystemScanner.scan(
    root: Path | str,
    allow_empty: Literal["allow", "deny"] = "deny",
    sort_entries: bool = False,
) -> FilesystemTree
```

Fluxo:

1. valida que a raiz nao e vazia, existe e e um diretorio;
2. resolve a raiz para path absoluto;
3. percorre a arvore top-down com `Path.walk()`;
4. materializa arquivos e pastas imediatos de cada nivel;
5. preserva niveis sem arquivos e diretorios vazios;
6. devolve um `FilesystemTree`.

`sort_entries=True` ordena `dirnames` in-place e os nomes de arquivos,
produzindo uma travessia deterministica. Erros de leitura do filesystem sao
propagados como `OSError`.

Exemplo:

```python
from core import LocalFileSystemScanner

tree = LocalFileSystemScanner().scan(
    "/dados/importacao",
    allow_empty="deny",
    sort_entries=True,
)

for level in tree.levels:
    print(level.path, level.total_files, level.total_size)
```

## 8. GraphClientManager

`GraphClientManager` cria e mantem:

- `ClientSecretCredential`;
- `GraphServiceClient`;
- scopes, por padrao `https://graph.microsoft.com/.default`.

Uso preferencial:

```python
async with GraphClientManager(credentials) as manager:
    service = SharePointService(manager)
```

O encerramento do contexto fecha a credencial assincrona. O
`SharePointService` nao assume a propriedade do manager.

## 9. SharePointService

### 9.1 Sites

```python
resolve_site(sharepoint_url: str) -> SharePointSite
```

Converte uma URL humana para a rota Graph e aceita as formas de resposta
observadas na operacao `sites.with_url(...)`.

### 9.2 Bibliotecas

```python
list_site_drives(
    site: SharePointSite,
    *,
    pagination: bool = True,
    max_pages: int | None = None,
) -> DocumentLibraryCollection

find_drive_by_name(
    site: SharePointSite,
    name: str,
) -> DocumentLibrary | None

get_default_drive(site: SharePointSite) -> DocumentLibrary
get_drive_by_id(drive_id: str) -> DocumentLibrary
```

`pagination=False` devolve somente a primeira pagina. `max_pages` limita a
quantidade de paginas processadas no cliente.

### 9.3 Itens

```python
get_drive_root(library: DocumentLibrary) -> SharePointItem

iter_children(
    library: DocumentLibrary,
    parent: SharePointItem,
) -> AsyncIterator[SharePointItemCollection]

list_children(
    library: DocumentLibrary,
    parent: SharePointItem,
    filter: Callable[[SharePointItem], bool] | None = None,
) -> SharePointItemCollection
```

`iter_children` entrega uma colecao por pagina. `list_children` consome esse
iterador e acumula todas as paginas, opcionalmente filtrando cada item.

Listagens especializadas:

```python
get_children_folder(...) -> SharePointItemCollection
get_children_file(...) -> SharePointItemCollection
```

Buscas de filhos imediatos:

```python
find_child_by_name(...) -> SharePointItem | None
find_child_by_id(...) -> SharePointItem | None
find_child_by_web_url(...) -> SharePointItem | None
find_folder_by_name(...) -> SharePointItem | None
find_file_by_name(...) -> SharePointItem | None
find_child_folder_by_id(...) -> SharePointItem | None
find_child_folder_web_url(...) -> SharePointItem | None
```

Essas buscas nao sao recursivas. Elas percorrem as paginas dos filhos do pai
informado e retornam `None` quando nao encontram correspondencia.

### 9.4 Pastas

```python
create_folder(
    library: DocumentLibrary,
    parent: SharePointItem,
    folder_name: str,
    conflict_behavior: ConflictBehavior = "fail",
) -> SharePointItem

ensure_remote_folder_path(
    library: DocumentLibrary,
    root: SharePointItem,
    folders_parts: Sequence[str],
    conflict_behavior: ConflictBehavior = "fail",
) -> SharePointItem
```

`ensure_remote_folder_path` percorre os fragmentos em ordem, reutiliza pastas
existentes e cria as ausentes. O retorno e a ultima pasta resolvida.

### 9.5 Upload individual

```python
upload(
    library: DocumentLibrary,
    parent: SharePointItem,
    local_file: LocalFile,
    conflict_behavior: ConflictBehavior = "fail",
) -> FileUploadResult
```

Fluxo atual:

- ate `250_000_000` bytes: le o arquivo em memoria e usa `PUT /content`;
- acima desse limite: cria uma upload session e usa `LargeFileUploadTask`;
- chunks grandes usam limite de `60 * 1024 * 1024` bytes;
- o retorno publico de ambos os fluxos e `FileUploadResult`.

## 10. Parsers, Builders e URLs

`core/parse.py` converte modelos e envelopes do SDK para modelos do Core e
traduz `ODataError`.

`core/builders.py` valida nomes e politicas de conflito e monta:

- o `DriveItem` usado na criacao de pasta;
- o `PreparedUpload` usado pelo upload pequeno.

`core/urls.py` valida URLs do Graph e constroi:

- a rota de resolucao do site;
- o fragmento de criacao por nome;
- a URL absoluta de criacao de conteudo no drive.

Esses helpers sao reexportados atualmente por `core`, mas ainda devem ser
tratados como infraestrutura interna ate a consolidacao da API publica.

## 11. Erros

Todas as excecoes semanticas derivam de `MSCronosError`. Grupos principais:

- configuracao e autenticacao;
- URL e validacao de entrada;
- caminho e arquivo local;
- site, biblioteca e item nao encontrado;
- permissao, conflito e resposta do Graph;
- criacao de item;
- upload pequeno, grande, sessao e chunks.

`parse_o_data_error()` extrai codigo, mensagem e contexto da operacao para
mapear erros OData para a hierarquia do Core. Excecoes semanticas ja criadas
devem ser relancadas sem perder o traceback original.

## 12. Inconsistencias Conhecidas

Estas pendencias foram observadas no codigo atual e nao devem ser confundidas
com funcionalidades implementadas:

1. `parse_local_file()` usa `path.suffix or None`, mas `LocalFile.extension`
   exige `str`; arquivos sem extensao quebram esse contrato.
2. `find_file_by_name()` ainda nao anota `library` e `parent`.
3. `LocalFile.rename_on_disk()` move fisicamente o arquivo, apesar de o nome
   sugerir que o model poderia apenas representar outro nome remoto.
4. `upload()` rejeita zero bytes mesmo quando `LocalFile.allow_empty` e
   `"allow"`; a politica de arquivos vazios ainda nao esta sincronizada com o
   uploader.
5. parsers, builders e helpers internos ainda sao reexportados em
   `core/__init__.py`.
6. nao existe suite automatizada de testes.
7. a lista completa de dependencias esta congelada em `requirements.txt`, sem
   metadados de pacote ou separacao entre dependencias diretas e transitivas.

## 13. Proxima Fase: Staging de Diretorios

Modelos planejados:

```text
StagingFile
StagingFolder
StagingFileCollection
StagingFolderCollection
StagingDirectoryLevel
StagingDirectoryLevelCollection
StagingFilesystemTree
```

Contrato esperado:

```text
LocalFileSystemScanner
  -> FilesystemTree
  -> builder de staging
  -> StagingFilesystemTree
  -> orquestrador de upload
  -> resultado agregado
```

Obrigacoes:

- calcular caminhos relativos sem expor paths absolutos no destino;
- validar nomes remotos e comportamento de conflito antes da transferencia;
- criar cada pasta pai antes dos filhos;
- preservar diretorios vazios;
- selecionar upload pequeno ou grande para cada arquivo;
- nao carregar a arvore inteira em bytes;
- registrar sucesso, falha e destino por item;
- definir comportamento de retomada e falha parcial;
- respeitar throttling e erros transitorios do Graph.

## 14. Criterios Para Beta

Uma beta utilizavel deve incluir:

- inconsistencias de assinatura corrigidas;
- staging local-remoto implementado;
- upload de uma arvore completa com diretorios vazios;
- resultado agregado e falha parcial documentados;
- testes unitarios dos models, scanner, parsers e staging;
- testes de paginacao e conflito;
- ao menos um teste de integracao controlado;
- API publica delimitada com `__all__`;
- metadados basicos de pacote e licenca definidos.
