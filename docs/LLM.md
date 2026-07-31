# LLM Repository Guide

Este arquivo oferece contexto operacional para assistentes que trabalham no
MS Cronos. O codigo e `docs/SPEC.md` continuam sendo as fontes de verdade.

## Resumo

MS Cronos e um wrapper Python assincrono para SharePoint via Microsoft Graph.
O pacote reduz a exposicao dos modelos do SDK por meio de modelos internos,
colecoes tipadas, parsers, builders e uma camada de servico.

O projeto possui dois fluxos:

```text
Fluxo remoto
GraphClientManager -> SharePointService -> SharePoint

Fluxo local
LocalFileSystemScanner -> FilesystemTree -> staging futuro -> upload futuro
```

Navegacao, criacao de pastas e uploads individuais estao implementados. O
scanner local tambem esta implementado. A associacao entre uma arvore local e
uma arvore remota ainda nao existe.

## Ambiente

- Python minimo: 3.12;
- SDK principal: `msgraph-sdk==1.58.0`;
- entrada publica recomendada: `from core import ...`;
- `main.py`: verificacao manual do scanner local;
- testes automatizados: ainda inexistentes.

Ao inventariar o repositorio, ignore:

```text
venv/
.venv/
__pycache__/
.pytest_cache/
.mypy_cache/
.ruff_cache/
*.pyc
```

## Ordem De Leitura

1. `README.md`: uso publico e status.
2. `docs/SPEC.md`: contratos, fluxos e pendencias conhecidas.
3. `core/models.py`: vocabulário do dominio.
4. `core/filesystem.py`: scanner do disco.
5. `core/sharepoint.py`: orquestracao remota.
6. `core/parse.py`: fronteira SDK -> Core.
7. `core/builders.py` e `core/urls.py`: payloads e rotas.
8. `core/errors.py`: taxonomia de falhas.
9. `core/graph_client.py`: autenticacao e ciclo de vida.
10. `core/__init__.py`: superficie atualmente reexportada.

## Mapa De Modulos

```text
core/
├── __init__.py       # reexports atuais
├── builders.py       # payload de pasta e politicas de conflito
├── errors.py         # erros semanticos
├── filesystem.py     # scanner local
├── graph_client.py   # credencial e GraphServiceClient
├── models.py         # modelos e colecoes
├── parse.py          # SDK/OData para Core
├── sharepoint.py     # servico remoto
├── urls.py           # validacao e montagem de URLs
└── utils.py          # rename_with_uuid
```

## Vocabulário Atual

Nao reintroduza os nomes antigos `SiteRef`, `DriveRef`, `DriveItemRef`,
`UploadFileResult` ou `StagingContentUpload`.

Use:

```text
SharePointSite
DocumentLibrary
SharePointItem
FileUploadResult
LocalFile
LocalFolder
DirectoryLevel
FilesystemTree
StagingFile
StagingFolder
StagingDirectoryLevel
StagingFilesystemTree
```

Colecoes publicas de retorno sao imutaveis e concretas:

```text
SharePointSiteCollection
DocumentLibraryCollection
SharePointItemCollection
LocalFileCollection
LocalFolderCollection
DirectoryLevelCollection
```

## Regras De Arquitetura

1. Modelos do Graph nao devem aparecer no retorno publico de
   `SharePointService`.
2. Conversoes do SDK pertencem a `core/parse.py`.
3. Montagem de payloads pertence a `core/builders.py`.
4. Montagem de URLs pertence a `core/urls.py`.
5. Erros OData sao traduzidos na fronteira do servico.
6. O scanner nao deve conhecer IDs, URLs ou conflitos do SharePoint.
7. Models de snapshot descrevem o disco; models de staging descrevem uma
   intencao de upload.
8. Nao leia bytes durante o scan. Abra o arquivo somente no momento do upload.
9. Preserve excecoes semanticas com `raise`; ao encadear outra falha, use
   `raise ... from error`.
10. Trabalhe com mudancas locais existentes; nao reverta alteracoes do usuario.

## API Remota

### Sites e bibliotecas

```python
site = await service.resolve_site(sharepoint_url)
libraries = await service.list_site_drives(site)
library = await service.get_default_drive(site)
library = await service.get_drive_by_id(drive_id)
library = await service.find_drive_by_name(site, name)
```

### Itens e paginacao

```python
root = await service.get_drive_root(library)

async for page in service.iter_children(library, root):
    for item in page:
        print(item.name)

all_items = await service.list_children(library, root)
folders = await service.get_children_folder(library, root)
files = await service.get_children_file(library, root)
```

`iter_children` e a primitiva paginada. `list_children` e o acumulador.
Buscas por nome, ID e URL percorrem filhos imediatos, nao descendentes
recursivos.

### Pastas e upload

```python
destination = await service.ensure_remote_folder_path(
    library,
    root,
    ("Financeiro", "2026"),
)

result = await service.upload(
    library,
    destination,
    LocalFile.from_uri("/tmp/relatorio.csv"),
    conflict_behavior="rename",
)
```

`upload()` seleciona:

- `PUT /content` ate `250_000_000` bytes;
- upload session acima desse limite;
- chunks de ate `60 * 1024 * 1024` bytes no fluxo grande.

O retorno sempre deve permanecer `FileUploadResult`.

## API Local

```python
from core import LocalFileSystemScanner

tree = LocalFileSystemScanner().scan(
    "/dados",
    allow_empty="deny",
    sort_entries=True,
)

for level in tree.levels:
    print(level.path, level.files, level.folders)
```

`FilesystemTree` e uma representacao plana. Nao presumir que cada nivel
armazena recursivamente seus descendentes. A ordem top-down e os caminhos
permitem reconstruir a hierarquia.

## Casos De Borda

### Scanner

- string vazia: `ValueError`;
- path inexistente: `ValueError`;
- path que nao e diretorio: `NotADirectoryError`;
- falha de permissao/leitura: `OSError`;
- diretorio vazio: gera um `DirectoryLevel`;
- arquivo vazio: rejeitado por padrao, aceito com `allow_empty="allow"`;
- arquivo sem extensao: `LocalFile.extension` deve ser `""`, nunca `None`;
- symlinks: seguir a semantica padrao de `Path.walk()` ate haver politica
  explicita no projeto.

### Paginacao

- uma pagina vazia nao implica necessariamente ausencia de proxima pagina;
- o `PageIterator` controla o proprio estado depois de `next()`;
- nao sincronizar manualmente `current_page`;
- `list_children` pode consumir muita memoria em bibliotecas grandes;
- para buscas, descarte paginas sem correspondencia em vez de acumular tudo.

### Upload

- o pai remoto deve ser pasta;
- o path local deve existir e ser arquivo legivel;
- `fail`, `rename` e `replace` devem conservar significado nos dois fluxos;
- URLs de upload session sao sensiveis e nao devem ser registradas;
- upload grande deve manter o stream aberto durante toda a task;
- resultados do SDK devem ser convertidos antes do retorno.

## Inconsistencias Que Nao Devem Ser Escondidas

1. `parse_local_file()` pode passar `None` como extensao.
2. `find_file_by_name()` possui dois parametros sem anotacao.
3. `LocalFile.rename_on_disk()` renomeia fisicamente o arquivo local.
4. a politica `allow_empty` do scanner nao e aceita pelo uploader atual.
5. helpers internos ainda sao reexportados por `core/__init__.py`.
6. nao ha testes automatizados.

Nao corrija esses pontos incidentalmente durante uma tarefa exclusivamente
documental. Registre ou trate cada um com escopo e teste proprios.

## Fase Seguinte

Implementar, nesta ordem:

1. modelos `StagingFile` e `StagingFolder`;
2. colecoes de staging;
3. `StagingDirectoryLevel`;
4. `StagingFilesystemTree`;
5. builder que calcula paths relativos e destinos;
6. orquestrador que cria pastas antes dos arquivos;
7. resultado agregado e politica de falha parcial;
8. retry/backoff para throttling e erros transitorios;
9. testes unitarios e integracao controlada.

O fluxo alvo e:

```text
FilesystemTree
  -> StagingFilesystemTree
  -> garantir pasta remota do nivel
  -> enviar arquivos do nivel
  -> prosseguir top-down
  -> produzir relatorio agregado
```

## Validacao Antes De Encerrar Mudancas

Para alteracoes de codigo:

```bash
python -m compileall core main.py
```

Quando os testes forem adicionados:

```bash
python -m pytest
```

Tambem confira:

- modelos do SDK nao vazaram para assinaturas publicas;
- imports continuam partindo de `core` para consumidores;
- docs usam os nomes atuais;
- `.env`, secrets, tokens e upload URLs nao aparecem no diff;
- venv e caches nao entram no inventario nem no commit.
