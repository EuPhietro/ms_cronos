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
