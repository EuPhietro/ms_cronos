# MS Cronos

MS Cronos e um wrapper Python assincrono para navegacao e upload de arquivos em
bibliotecas de documentos do SharePoint por meio do Microsoft Graph.

O projeto oferece modelos de dominio pequenos e uma API de alto nivel para que
aplicacoes consumidoras nao precisem lidar diretamente com `DriveItem`,
envelopes OData, URLs de continuacao ou detalhes de upload do SDK.

> Status: desenvolvimento ativo. A navegacao remota e os uploads de arquivos
> estao funcionais, mas a API ainda pode mudar antes da primeira versao estavel.

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
- traducao de erros OData para excecoes do Core.

O objetivo seguinte e enviar arvores completas de diretorios locais,
preservando sua estrutura no SharePoint.

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
resultado de upload   ->     FileUploadResult
```

As operacoes de `SharePointService` convertem objetos e envelopes do SDK antes
de devolve-los ao consumidor.

## Requisitos

- Python 3.11 ou superior;
- uma aplicacao registrada no Microsoft Entra ID;
- permissoes de aplicacao adequadas para os sites e arquivos acessados;
- consentimento administrativo quando exigido pelo tenant.

## Instalacao

A partir de um checkout do repositorio:

```bash
python -m venv venv
source venv/bin/activate
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
```

No PowerShell:

```powershell
python -m venv venv
.\venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
```

O projeto ainda nao possui uma distribuicao publicada no PyPI.

## Configuracao

Crie um arquivo `.env` local:

```env
CLIENT_ID=identificador-da-aplicacao
CLIENT_SECRET=segredo-da-aplicacao
CLIENT_TENANT=identificador-do-tenant
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

## Navegacao Remota

### Resolver um site

```python
site = await sharepoint.resolve_site(
    "https://tenant.sharepoint.com/sites/Financeiro"
)

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
- `LocalFileCollection`.

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

`main.py` funciona como laboratorio de integracao contra um tenant real:

```bash
python main.py
```

O repositorio ainda nao possui uma suite automatizada de testes. Antes de uma
versao estavel, o projeto precisa de testes unitarios para parsers, paginacao,
conflitos e uploads, alem de testes de integracao opcionais.

## Roadmap

- consolidar a superficie publica com `__all__`;
- manter parsers e builders fora dos exports publicos;
- modelar diretorios e arvores locais;
- implementar upload recursivo de diretorios;
- preservar diretorios vazios;
- adicionar resultados agregados e relatorios de falha parcial;
- implementar retries para throttling e erros transitorios;
- publicar testes automatizados;
- preparar empacotamento e versionamento.

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

Este repositorio ainda nao inclui um arquivo de licenca. Antes de reutilizar ou
redistribuir o projeto, aguarde a publicacao dos termos aplicaveis.
