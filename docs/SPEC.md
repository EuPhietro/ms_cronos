# MS Cronos Specification

Este documento descreve a especificacao geral do MS Cronos.

O projeto deve ser pequeno, modular e escrito em Python. A primeira entrega concreta e um Core capaz de fazer upload de arquivos locais para uma biblioteca de documentos do SharePoint usando Microsoft Graph. Depois disso, o Core sera usado por uma API simples de jobs e por um daemon Windows.

## Visao Geral

MS Cronos deve automatizar o envio de arquivos locais para o SharePoint.

O fluxo final esperado e:

```text
API ou configuracao local
  -> cria job de upload
  -> daemon Windows busca job pendente
  -> daemon chama Core
  -> Core envia arquivo para SharePoint
  -> job recebe status final
```

O projeto deve evitar acoplamento entre as camadas. O Core nao deve saber que existe daemon. O daemon nao deve saber detalhes do Microsoft Graph. A API nao deve executar upload diretamente; ela deve apenas criar, consultar e eventualmente cancelar jobs.

## Objetivos

- Criar um wrapper Python simples para Microsoft Graph e SharePoint.
- Facilitar upload de arquivos para bibliotecas de documentos.
- Permitir que jobs sejam criados por API no futuro.
- Permitir execucao em background no Windows.
- Manter o Core pequeno e testavel.
- Evitar que regras de negocio fiquem misturadas com chamadas do SDK.

## Nao Objetivos

- Criar UI.
- Criar sincronizacao bidirecional.
- Criar monitoramento automatico de diretorios na primeira versao.
- Suportar upload grande na primeira versao.
- Usar banco complexo na primeira versao.
- Fazer abstracao generica para todos os recursos do Microsoft Graph.
- Expor diretamente todos os modelos do `msgraph-sdk`.

## Camadas

O projeto deve ser dividido nestas camadas:

```text
Presentation / API
  -> recebe comandos externos e cria jobs

Application / Jobs
  -> define contratos de execucao, status e validacoes de job

Daemon / Worker
  -> executa jobs em background no Windows

Core
  -> implementa operacoes SharePoint/Drive usando Microsoft Graph

Infrastructure
  -> configuracao, credenciais, logging, persistencia e empacotamento
```

Cada camada deve depender apenas das camadas abaixo dela.

## Estrutura Planejada

Estrutura alvo sugerida:

```text
ms_cronos/
  __init__.py

  core/
    __init__.py
    graph_client.py
    sharepoint.py
    models.py
    errors.py
    urls.py

  jobs/
    __init__.py
    models.py
    repository.py
    service.py
    errors.py

  daemon/
    __init__.py
    worker.py
    service.py
    windows_service.py

  api/
    __init__.py
    app.py
    routes.py
    schemas.py

  infra/
    __init__.py
    config.py
    logging.py
    storage.py

main.py
README.md
SPEC.md
docs/
  CORE_SPEC.md
requirements.txt
```

Durante a fase inicial, a estrutura pode ser menor. O importante e preservar as responsabilidades.

## Camada Core

O Core e a camada mais importante da primeira fase.

Responsabilidades:

- criar cliente autenticado do Microsoft Graph;
- resolver site por URL do SharePoint;
- obter drive padrao de um site;
- obter raiz do drive;
- listar filhos de uma pasta;
- encontrar arquivo ou pasta por nome;
- criar pasta;
- enviar arquivo pequeno;
- sobrescrever arquivo existente;
- converter modelos do SDK para dataclasses internas;
- relancar erros do Graph como erros do projeto.

Fora do Core:

- ler `.env`;
- imprimir no terminal;
- gravar jobs;
- rodar loop infinito;
- expor endpoint HTTP;
- saber que existe daemon Windows.

### Modulos do Core

`core/graph_client.py`

- Deve construir `GraphServiceClient`.
- Deve receber credenciais por objeto.
- Deve cuidar do fechamento do `ClientSecretCredential`.
- Deve oferecer factory ou context manager assincrono.
- Nao deve ler variaveis de ambiente diretamente.

`core/sharepoint.py`

- Deve conter `SharePointDriveClient`.
- Deve concentrar chamadas para `sites`, `drives`, `items`, `children` e `content`.
- Deve esconder a verbosidade do SDK.
- Deve retornar modelos proprios do projeto.

`core/models.py`

- Deve conter dataclasses leves e preferencialmente imutaveis.
- Deve conter modelos de entrada e saida.
- Nao deve importar `msgraph`.

`core/errors.py`

- Deve conter excecoes do projeto.
- Deve preservar contexto de falhas.
- Deve evitar que `ODataError` vaze para as camadas superiores.

`core/urls.py`

- Deve conter funcoes pequenas para montar URLs Graph quando o SDK nao tiver builder adequado.
- Deve escapar nomes de arquivo e caminhos com seguranca.
- Deve ser coberto por testes unitarios.

### API Publica do Core

Interface alvo:

```python
class SharePointDriveClient:
    async def close(self) -> None: ...

    async def resolve_site(self, sharepoint_url: str) -> SiteRef: ...
    async def get_default_drive(self, site_id: str) -> DriveRef: ...
    async def get_drive_root(self, drive_id: str) -> DriveItemRef: ...

    async def list_children(
        self,
        drive_id: str,
        folder_id: str,
    ) -> list[DriveItemRef]: ...

    async def find_child(
        self,
        drive_id: str,
        folder_id: str,
        name: str,
    ) -> DriveItemRef | None: ...

    async def find_child_folder(
        self,
        drive_id: str,
        folder_id: str,
        name: str,
    ) -> DriveItemRef | None: ...

    async def create_folder(
        self,
        drive_id: str,
        parent_id: str,
        name: str,
        conflict_behavior: str = "rename",
    ) -> DriveItemRef: ...

    async def upload_file(
        self,
        drive_id: str,
        parent_id: str,
        local_path: str,
        remote_name: str | None = None,
        conflict_behavior: str = "replace",
    ) -> UploadResult: ...
```

### Modelos do Core

`GraphCredentials`

- `tenant_id: str`
- `client_id: str`
- `client_secret: str`

`SiteRef`

- `id: str`
- `name: str | None`
- `display_name: str | None`
- `web_url: str | None`

`DriveRef`

- `id: str`
- `name: str | None`
- `web_url: str | None`
- `drive_type: str | None`

`DriveItemRef`

- `id: str`
- `name: str | None`
- `web_url: str | None`
- `is_folder: bool`
- `is_file: bool`
- `size: int | None`

`UploadResult`

- `item: DriveItemRef`
- `source_path: Path`
- `remote_name: str`
- `conflict_behavior: str`

### Objetivo Operacional do Core

O Core deve permitir que outra camada realize este fluxo sem conhecer o SDK:

```text
recebe credenciais
  -> cria cliente Graph
  -> resolve site por URL
  -> encontra drive padrao
  -> encontra pasta de destino
  -> cria pasta se necessario
  -> envia arquivo
  -> retorna resultado estavel
```

O chamador deve conseguir fazer upload para SharePoint trabalhando apenas com:

- credenciais;
- URL do site SharePoint;
- nome da pasta;
- caminho local do arquivo;
- politica de conflito.

### Limites Funcionais da Primeira Versao

O Core v1 deve suportar:

- apenas SharePoint via Microsoft Graph;
- apenas autenticacao por credenciais de aplicacao;
- apenas drive padrao do site;
- apenas upload simples de arquivo pequeno;
- apenas operacoes em uma pasta por vez;
- apenas nomes simples de pasta e arquivo;
- apenas execucao assincrona.

O Core v1 nao deve suportar:

- upload resumivel;
- upload paralelo em chunks;
- multiplos drives por alias funcional;
- descoberta automatica de site por busca textual;
- regras de sincronizacao;
- exclusao de arquivos;
- renomeacao de arquivos ou pastas;
- mover arquivos entre pastas;
- observabilidade baseada em eventos externos.

### Contratos de Entrada

Entradas aceitas pelo Core:

`GraphCredentials`

- `tenant_id` obrigatorio, nao vazio.
- `client_id` obrigatorio, nao vazio.
- `client_secret` obrigatorio, nao vazio.

`sharepoint_url`

- deve ser URL absoluta;
- deve usar `https`;
- deve conter hostname;
- deve apontar para um site SharePoint;
- nao deve incluir sufixos de recurso Graph como `:/drive`.

`site_id`

- deve ser string nao vazia;
- deve ser o identificador retornado pelo Graph;
- o Core nao deve tentar adivinhar ou corrigir `site_id`.

`drive_id`

- deve ser string nao vazia;
- deve ser um identificador de drive, nao de site.

`folder_id`

- deve ser string nao vazia;
- deve representar um `DriveItem` do tipo pasta.

`local_path`

- deve existir no disco;
- deve ser arquivo regular;
- deve ser legivel pelo processo;
- nao deve exceder o limite de upload simples do v1.

`remote_name`

- quando informado, deve ser string nao vazia;
- nao deve conter barras ou segmentos de caminho;
- quando omitido, o nome do arquivo local deve ser usado.

`conflict_behavior`

- valores aceitos no v1: `replace`, `rename`, `fail`;
- qualquer outro valor deve levantar erro antes da chamada ao Graph.

### Contratos de Saida

O Core deve sempre retornar modelos proprios do projeto.

Regras:

- nunca retornar `Drive`, `Site`, `DriveItem` ou `ODataError` diretamente;
- nunca retornar `dict` cru do SDK como contrato publico;
- nunca depender de `additional_data` fora da camada de adaptacao.

Saidas por metodo:

- `resolve_site` retorna `SiteRef`.
- `get_default_drive` retorna `DriveRef`.
- `get_drive_root` retorna `DriveItemRef`.
- `list_children` retorna `list[DriveItemRef]`.
- `find_child` retorna `DriveItemRef | None`.
- `find_child_folder` retorna `DriveItemRef | None`.
- `create_folder` retorna `DriveItemRef`.
- `upload_file` retorna `UploadResult`.

### Comportamento por Metodo

`resolve_site(sharepoint_url: str) -> SiteRef`

- valida a URL recebida;
- converte a URL humana do SharePoint para a rota Graph de site;
- consulta o Graph;
- extrai o identificador do site;
- retorna `SiteRef`.

Falhas esperadas:

- URL invalida;
- URL sem hostname;
- resposta sem `site_id`;
- permissao insuficiente;
- site inexistente.

`get_default_drive(site_id: str) -> DriveRef`

- valida `site_id`;
- consulta `/sites/{site-id}/drive`;
- retorna o drive padrao do site;
- nao deve listar outros drives no v1.

Falhas esperadas:

- `site_id` vazio;
- site sem drive padrao acessivel;
- permissao insuficiente.

`get_drive_root(drive_id: str) -> DriveItemRef`

- valida `drive_id`;
- consulta `/drives/{drive-id}/root`;
- garante que o item retornado e pasta;
- retorna `DriveItemRef`.

Falhas esperadas:

- drive inexistente;
- drive inacessivel;
- resposta sem `id`.

`list_children(drive_id: str, folder_id: str) -> list[DriveItemRef]`

- valida `drive_id` e `folder_id`;
- consulta os filhos da pasta;
- converte a colecao para `DriveItemRef`;
- retorna lista vazia quando a pasta nao tem filhos.

Falhas esperadas:

- pasta nao encontrada;
- item informado nao e pasta;
- permissao insuficiente.

`find_child(drive_id: str, folder_id: str, name: str) -> DriveItemRef | None`

- chama `list_children`;
- compara por nome exato;
- retorna o primeiro item encontrado;
- retorna `None` quando nao existir.

Regras:

- comparacao inicial deve ser exata;
- no v1 nao deve haver matching parcial nem regex;
- no v1 a sensibilidade de caixa deve seguir o nome retornado pelo Graph e ser documentada como comparacao exata.

`find_child_folder(drive_id: str, folder_id: str, name: str) -> DriveItemRef | None`

- chama `find_child`;
- filtra apenas itens do tipo pasta;
- retorna `None` se o nome existir mas for arquivo.

`create_folder(drive_id: str, parent_id: str, name: str, conflict_behavior: str = "rename") -> DriveItemRef`

- valida `drive_id`, `parent_id`, `name` e `conflict_behavior`;
- cria um `DriveItem` com facet `Folder`;
- envia `@microsoft.graph.conflictBehavior`;
- retorna a pasta criada.

Falhas esperadas:

- item pai nao e pasta;
- nome invalido;
- conflito nao permitido pelo comportamento escolhido.

`upload_file(...) -> UploadResult`

- valida argumentos;
- resolve o nome remoto;
- le bytes do arquivo local;
- monta a URL de upload por caminho;
- executa upload simples;
- converte resposta para `UploadResult`.

Regras:

- deve sobrescrever via `PUT /content` quando o comportamento for `replace`;
- deve criar novo nome quando o comportamento for `rename`, desde que a rota suportada preserve esse comportamento;
- deve falhar cedo quando `local_path` nao for arquivo valido;
- deve falhar cedo quando o tamanho exceder o limite do v1.

### Adaptacao entre SDK e Modelos do Projeto

O Core deve ter uma camada interna clara de conversao.

Recomendacao:

- uma funcao para converter `Site` ou resposta de resolucao para `SiteRef`;
- uma funcao para converter `Drive` para `DriveRef`;
- uma funcao para converter `DriveItem` para `DriveItemRef`;
- uma funcao para converter resposta de upload para `UploadResult`.

Regras de conversao:

- `is_folder` e `True` quando `item.folder` existir;
- `is_file` e `True` quando `item.file` existir;
- `size` pode ser `None`;
- `name` e `web_url` podem ser `None`, mas `id` nao.

### Estrutura Interna Recomendada do Core

Mesmo que o projeto ainda esteja pequeno, o Core deve ser separado em partes simples.

`core/graph_client.py`

- factory de `GraphServiceClient`;
- gerenciamento de credenciais;
- contexto de abertura e fechamento;
- nenhuma regra de SharePoint.

`core/sharepoint.py`

- API publica do Core;
- chamadas de fluxo de negocio tecnico;
- coordenacao entre validacoes, SDK e conversores.

`core/models.py`

- dataclasses publicas;
- tipos de retorno;
- nenhum acesso ao Graph.

`core/errors.py`

- hierarquia de erros;
- mensagens canônicas do Core;
- nenhum import de `rich`.

`core/urls.py`

- `build_graph_site_url(sharepoint_url: str) -> str`
- `build_upload_url(drive_id: str, parent_id: str, remote_name: str) -> str`
- funcoes puras e testaveis.

`core/validators.py`

- validacao de URL;
- validacao de `conflict_behavior`;
- validacao de caminho local;
- validacao de tamanho de arquivo;
- validacao de nome remoto.

`core/adapters.py`

- adaptacao de modelos do SDK para modelos do projeto;
- nenhuma regra de negocio fora da traducao.

### Sequencia Interna das Operacoes

Fluxo interno recomendado para upload:

```text
upload_file
  -> validate_local_path
  -> validate_conflict_behavior
  -> resolve_remote_name
  -> read_file_bytes
  -> build_upload_url
  -> execute_put_content
  -> adapt_drive_item_to_upload_result
```

Fluxo interno recomendado para criar pasta:

```text
create_folder
  -> validate_name
  -> validate_conflict_behavior
  -> build_folder_drive_item
  -> execute_children_post
  -> adapt_drive_item_to_ref
```

### Validacoes Obrigatorias

O Core deve validar antes de chamar o Graph:

- credenciais nao vazias;
- URL SharePoint bem formada;
- `site_id`, `drive_id`, `parent_id`, `folder_id` nao vazios;
- nome de pasta nao vazio;
- nome remoto nao vazio quando fornecido;
- `conflict_behavior` valido;
- caminho local existente;
- caminho local apontando para arquivo;
- tamanho de arquivo dentro do limite suportado.

Validacoes de nome devem ser conservadoras no v1:

- bloquear nome vazio;
- bloquear barra `/`;
- bloquear barra invertida `\\`;
- bloquear nomes com espacos apenas.

### Politica de Erros do Core

O Core deve traduzir erros do Graph para erros de dominio tecnico do projeto.

Mapeamento sugerido:

- URL invalida -> `SharePointUrlError`
- credenciais invalidas -> `GraphConfigurationError`
- falha de permissao -> `GraphRequestError`
- item nao encontrado -> `DriveItemNotFoundError`
- item esperado como pasta mas recebido como arquivo -> `NotAFolderError`
- arquivo local inexistente ou inacessivel -> `LocalFileError`
- conflito invalido -> `UnsupportedConflictBehaviorError`
- arquivo acima do limite -> `LargeFileUploadNotSupportedError`

`GraphRequestError` deve preservar:

- operacao;
- recurso alvo;
- codigo do Graph, quando existir;
- mensagem do Graph, quando existir.

### Politica de Logging no Core

O Core nao deve imprimir no terminal.

Regras:

- nao usar `print`;
- nao usar `rich`;
- nao formatar traceback para o usuario final;
- usar `logging.getLogger(__name__)` apenas quando necessario;
- nao logar segredo, token ou caminho sensivel sem necessidade.

No v1, o Core pode operar sem logs detalhados se isso simplificar a implementacao, desde que os erros retornem contexto suficiente.

### Politica de Dependencias

O Core deve depender apenas do necessario:

- `azure-identity`
- `msgraph-sdk`
- biblioteca padrao

Dependencias que nao pertencem ao Core:

- `python-dotenv`
- `rich`
- frameworks web
- bibliotecas de servico Windows

### Compatibilidade e Evolucao

O Core deve ser desenhado para crescer sem quebrar o contrato do v1.

Isso significa:

- manter a API publica pequena;
- evitar expor detalhes do SDK;
- centralizar montagem de URLs;
- centralizar adaptacao de modelos;
- deixar espaco para `createUploadSession` no futuro sem mudar o chamador.

### Criterios de Aceite do Core

O Core estara pronto para a primeira entrega quando:

1. O laboratorio `main.py` usar o Core em vez de chamar o SDK diretamente.
2. `resolve_site` retornar `SiteRef` valido.
3. `get_default_drive` retornar `DriveRef` valido.
4. `get_drive_root` retornar pasta raiz como `DriveItemRef`.
5. `list_children` funcionar para raiz e subpastas.
6. `find_child_folder` localizar pasta existente.
7. `create_folder` criar pasta com comportamento de conflito configuravel.
8. `upload_file` enviar arquivo pequeno com sucesso.
9. Nenhum metodo publico do Core retornar objetos do SDK.
10. O Core nao usar `.env`, `print` ou `rich`.

### Regras do Core

- O Core deve ser assincrono.
- O Core deve aceitar um `GraphServiceClient` injetado para testes.
- O Core deve validar entradas antes de chamar o Graph.
- O Core deve retornar objetos estaveis do projeto.
- O Core deve encapsular URLs especiais de upload.
- O Core deve ter mensagens de erro claras.
- O Core deve limitar upload simples a arquivos pequenos.

### Rotas Graph Usadas

```http
GET /sites/{hostname}:/{site-path}
GET /sites/{site-id}/drive
GET /drives/{drive-id}/root
GET /drives/{drive-id}/items/{folder-id}/children
POST /drives/{drive-id}/items/{folder-id}/children
PUT /drives/{drive-id}/items/{folder-id}:/{filename}:/content
PUT /drives/{drive-id}/items/{file-id}/content
```

Essas rotas nao devem aparecer fora do Core.

## Camada Jobs

A camada de jobs representa o contrato de execucao. Ela nao deve executar upload diretamente. Ela deve descrever o que precisa ser feito.

Responsabilidades:

- definir modelo de job;
- validar entrada de job;
- controlar status;
- armazenar tentativas;
- registrar erro de execucao;
- fornecer jobs pendentes para o daemon.

### Modelo de Job

Campos iniciais sugeridos:

```python
class UploadJob:
    id: str
    status: JobStatus
    sharepoint_url: str
    target_folder: str
    local_path: str
    drive: str = "default"
    remote_name: str | None = None
    conflict_behavior: str = "rename"
    attempts: int = 0
    max_attempts: int = 3
    created_at: datetime
    updated_at: datetime
    started_at: datetime | None = None
    finished_at: datetime | None = None
    last_error: str | None = None
```

### Status de Job

Status iniciais:

- `pending`: criado, aguardando execucao.
- `running`: daemon iniciou execucao.
- `done`: upload concluido.
- `failed`: falha definitiva.
- `retrying`: falha recuperavel, aguardando nova tentativa.
- `cancelled`: cancelado antes de concluir.

Transicoes permitidas:

```text
pending -> running
running -> done
running -> retrying
running -> failed
running -> cancelled
retrying -> running
pending -> cancelled
```

### Regras de Job

- Um job deve ser imutavel em seus campos principais depois de iniciar.
- Um job em `done` nao deve voltar para `pending`.
- Um job em `failed` pode ser clonado ou reaberto apenas por decisao explicita da API.
- Tentativas devem ser incrementadas pelo daemon, nao pela API.
- Erros devem preservar mensagem curta e contexto tecnico.

## Camada Daemon

O daemon sera responsavel por executar jobs no Windows.

Responsabilidades:

- iniciar como processo de background;
- carregar configuracao;
- criar instancia do Core;
- buscar jobs pendentes;
- executar uploads;
- atualizar status de jobs;
- aplicar retry;
- registrar logs;
- encerrar com seguranca.

Fora do daemon:

- implementar regras do Microsoft Graph;
- montar URLs Graph;
- validar detalhes internos de SharePoint;
- expor API HTTP.

### Loop do Daemon

Fluxo esperado:

```text
start
  -> load config
  -> create graph/core client
  -> while running:
       -> fetch pending jobs
       -> lock one job
       -> mark running
       -> execute upload through Core
       -> mark done or retrying/failed
       -> sleep interval
  -> close clients
stop
```

### Configuracoes do Daemon

Campos sugeridos:

- `poll_interval_seconds`
- `max_parallel_jobs`
- `max_attempts`
- `retry_delay_seconds`
- `log_level`
- `storage_path`
- `service_name`

Na primeira versao, `max_parallel_jobs` pode ser `1`.

### Instalacao como Servico Windows

A camada Windows deve ficar isolada.

Opcoes futuras:

- `pywin32`
- `nssm`
- Windows Task Scheduler

Para a primeira versao, o daemon pode rodar como comando Python comum antes de virar servico.

## Camada API

A API sera uma camada posterior para criar e consultar jobs.

Responsabilidades:

- receber requisicoes HTTP;
- validar payload;
- criar jobs;
- consultar jobs;
- cancelar jobs pendentes;
- nao executar upload diretamente.

Endpoints sugeridos:

```http
POST /jobs
GET /jobs
GET /jobs/{job_id}
POST /jobs/{job_id}/cancel
POST /jobs/{job_id}/retry
```

Payload inicial para criar job:

```json
{
  "sharepoint_url": "https://tenant.sharepoint.com/sites/RHConecta",
  "target_folder": "Curriculos",
  "local_path": "C:/arquivos/curriculo.pdf",
  "drive": "default",
  "remote_name": null,
  "conflict_behavior": "rename"
}
```

Resposta esperada:

```json
{
  "id": "job_123",
  "status": "pending",
  "created_at": "2026-05-29T13:00:00Z"
}
```

## Camada Infrastructure

Infrastructure contem adaptadores tecnicos que nao pertencem ao dominio.

Responsabilidades:

- carregar configuracao;
- configurar logs;
- persistir jobs;
- guardar estado local;
- empacotar aplicacao;
- integrar com Windows.

### Configuracao

Fontes permitidas:

- variaveis de ambiente;
- `.env` para desenvolvimento;
- arquivo `.toml` ou `.json` para instalacao local;
- parametros de linha de comando para scripts.

O Core nao deve carregar configuracao sozinho.

### Persistencia

Primeira versao sugerida:

- SQLite local.

Motivos:

- simples;
- embutido no Python;
- suficiente para jobs locais;
- bom para daemon Windows pequeno.

Tabelas iniciais:

```text
jobs
  id
  status
  sharepoint_url
  drive
  target_folder
  local_path
  remote_name
  conflict_behavior
  attempts
  max_attempts
  created_at
  updated_at
  started_at
  finished_at
  last_error
```

### Logging

Logs devem ser estruturados o suficiente para operacao.

Campos recomendados:

- timestamp;
- level;
- component;
- job_id;
- operation;
- message;
- error_code;
- exception_type.

O Core deve usar `logging`, nao `print`.

`rich` pode ser usado apenas no laboratorio, CLI ou debug local.

## Erros

Erros devem ser separados por camada.

Core:

- `GraphConfigurationError`
- `SharePointUrlError`
- `GraphRequestError`
- `DriveItemNotFoundError`
- `NotAFolderError`
- `LocalFileError`
- `UnsupportedConflictBehaviorError`
- `LargeFileUploadNotSupportedError`

Jobs:

- `JobValidationError`
- `JobNotFoundError`
- `InvalidJobTransitionError`
- `JobAlreadyFinishedError`

Daemon:

- `DaemonConfigurationError`
- `JobExecutionError`
- `WorkerShutdownError`

API:

- Deve traduzir erros em HTTP status codes.
- Nao deve retornar traceback cru.

## Seguranca

- Nunca versionar `.env`.
- Nunca logar `CLIENT_SECRET`.
- Nunca retornar secrets pela API.
- Validar caminhos locais antes de executar jobs.
- Considerar allowlist de diretorios locais no daemon.
- Preferir permissoes Graph minimas.
- Avaliar `Sites.Selected` quando possivel.

## Upload

### Upload Pequeno

Primeira versao:

- suportar arquivo pequeno;
- limite sugerido: 4 MB;
- usar `PUT /content`;
- retornar `UploadResult`.

### Upload Grande

Fase posterior:

- usar `createUploadSession`;
- dividir arquivo em chunks;
- permitir retry por chunk;
- registrar progresso.

## Testes

Tipos de teste:

- unitarios para Core;
- unitarios para jobs;
- unitarios para montagem de URLs;
- unitarios para validacoes locais;
- testes de integracao manuais contra SharePoint;
- teste end-to-end futuro com API + daemon + Core.

Casos minimos do Core:

- resolver site por URL valida;
- rejeitar URL invalida;
- listar filhos;
- encontrar pasta existente;
- retornar `None` para pasta inexistente;
- criar pasta;
- validar `conflict_behavior`;
- rejeitar arquivo inexistente;
- rejeitar diretorio em vez de arquivo;
- rejeitar arquivo grande;
- montar URL de upload com nome escapado.

Casos minimos de Jobs:

- criar job valido;
- rejeitar job sem caminho local;
- rejeitar transicao invalida;
- marcar `running`;
- marcar `done`;
- marcar `retrying`;
- marcar `failed`.

## Fases de Entrega

### Fase 1: Core Basico

Entregaveis:

- pacote Python do Core;
- modelos proprios;
- erros proprios;
- wrapper `SharePointDriveClient`;
- upload pequeno;
- script manual usando o Core.

Criterios de aceite:

- `main.py` nao chama mais o SDK diretamente;
- um arquivo pequeno pode ser enviado para uma pasta existente;
- uma pasta pode ser criada;
- erros sao legiveis.

### Fase 2: Jobs Locais

Entregaveis:

- modelo de job;
- repositorio SQLite;
- servico de jobs;
- CLI ou script para criar job manualmente.

Criterios de aceite:

- job pode ser criado;
- job pode ser consultado;
- status muda corretamente;
- falhas ficam registradas.

### Fase 3: Daemon

Entregaveis:

- worker que processa jobs pendentes;
- controle de retry;
- logs;
- encerramento seguro.

Criterios de aceite:

- daemon processa job pendente;
- daemon atualiza status;
- daemon nao perde erro;
- daemon consegue rodar continuamente.

### Fase 4: API

Entregaveis:

- API HTTP simples;
- endpoints de jobs;
- schemas de entrada e saida;
- traducao de erros para HTTP.

Criterios de aceite:

- API cria jobs;
- API lista jobs;
- API consulta job por ID;
- API cancela job pendente.

### Fase 5: Windows Service

Entregaveis:

- empacotamento;
- instalacao como servico ou task;
- configuracao local;
- documentacao operacional.

## Decisoes Abertas

- Nome final do pacote: `core` ou `ms_cronos.core`.
- Framework da API: FastAPI ou outro.
- Persistencia inicial: SQLite ou arquivo JSON.
- Metodo de instalacao no Windows: service, NSSM ou Task Scheduler.
- Limite exato para upload simples.
- Politica de retry.
- Modelo de configuracao para producao.

## Decisoes Recomendadas

- Usar `ms_cronos.core` como pacote final para evitar nome generico.
- Usar SQLite para jobs locais.
- Usar FastAPI apenas quando a camada API for iniciada.
- Usar upload simples no Core v1.
- Deixar upload grande para v2.
- Manter o Core sem `rich`, `.env` e `print`.
- Usar `logging` em todas as camadas produtivas.

## Glossario

`Site`

Representa um site do SharePoint.

`Drive`

Representa uma biblioteca de documentos.

`DriveItem`

Representa arquivo ou pasta dentro de um drive.

`Root`

Pasta raiz de uma biblioteca de documentos.

`Children`

Itens dentro de uma pasta.

`Job`

Pedido de execucao registrado para o daemon.

`Daemon`

Processo em background que executa jobs.

`Core`

Camada reutilizavel que conversa com Microsoft Graph.
