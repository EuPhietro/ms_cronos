# Core Specification

Esta especificacao define a primeira versao do Core do MS Cronos.

O Core e a camada Python responsavel por conversar com o Microsoft Graph e oferecer uma interface pequena para operacoes de SharePoint/Drive. Ele nao deve conhecer daemon Windows, API HTTP, banco de dados, filas ou agenda de execucao.

## Objetivo

Construir um wrapper leve sobre `msgraph-sdk` para permitir que outras camadas facam upload de arquivos para o SharePoint sem precisar lidar diretamente com a API fluente do SDK.

O Core deve resolver estes casos iniciais:

- autenticar no Microsoft Graph usando credenciais de aplicacao;
- resolver um site SharePoint por URL;
- obter a biblioteca de documentos padrao de um site;
- obter a raiz de um drive;
- listar itens de uma pasta;
- encontrar pasta por nome;
- criar pasta;
- enviar arquivo local para uma pasta;
- sobrescrever arquivo existente quando solicitado.

## Fora do Escopo

Nesta primeira versao, o Core nao deve implementar:

- daemon Windows;
- API HTTP;
- persistencia de jobs;
- scheduler;
- upload resumivel para arquivos grandes;
- sincronizacao bidirecional;
- monitoramento de diretorios;
- UI;
- regras de negocio especificas de RH, financeiro ou qualquer departamento.

Essas responsabilidades devem ficar em camadas futuras.

## Dependencias

Dependencias diretas esperadas:

- `azure-identity`
- `msgraph-sdk`
- `python-dotenv` apenas para scripts locais ou exemplos
- `rich` apenas para laboratorio, CLI e debug

O Core em si nao deve depender de `rich`. Saidas bonitas devem ficar em scripts, CLI ou camada de apresentacao.

## Modulos Planejados

```text
ms_cronos/
  __init__.py
  graph_client.py
  sharepoint.py
  models.py
  errors.py
```

Responsabilidades:

- `graph_client.py`: cria e fecha `GraphServiceClient`.
- `sharepoint.py`: contem `SharePointDriveClient`.
- `models.py`: dataclasses simples usadas como entrada e saida do Core.
- `errors.py`: excecoes especificas do projeto.

## Configuracao

O Core deve receber configuracao por objeto Python, nao ler `.env` diretamente.

```python
from dataclasses import dataclass

@dataclass(frozen=True)
class GraphCredentials:
    tenant_id: str
    client_id: str
    client_secret: str
```

Scripts e aplicacoes externas podem carregar `.env`, variaveis de ambiente, arquivo JSON ou outro mecanismo. Depois disso, passam `GraphCredentials` para o Core.

## Modelos Publicos

O Core deve retornar modelos pequenos, estaveis e independentes do SDK sempre que possivel.

```python
from dataclasses import dataclass
from pathlib import Path

@dataclass(frozen=True)
class SiteRef:
    id: str
    name: str | None
    display_name: str | None
    web_url: str | None

@dataclass(frozen=True)
class DriveRef:
    id: str
    name: str | None
    web_url: str | None
    drive_type: str | None

@dataclass(frozen=True)
class DriveItemRef:
    id: str
    name: str | None
    web_url: str | None
    is_folder: bool
    is_file: bool
    size: int | None

@dataclass(frozen=True)
class UploadResult:
    item: DriveItemRef
    source_path: Path
    conflict_behavior: str
```

Motivo: o restante do sistema nao deve depender diretamente de `DriveItem`, `Drive`, `Site` ou outros modelos gerados pelo Graph SDK.

## API Publica

Interface desejada:

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
        local_path: str | Path,
        remote_name: str | None = None,
        conflict_behavior: str = "replace",
    ) -> UploadResult: ...
```

## Comportamento de Upload

`upload_file` deve receber um caminho local e enviar o arquivo para uma pasta do SharePoint.

Comportamento esperado:

1. Validar se `local_path` existe.
2. Validar se `local_path` e arquivo.
3. Usar `remote_name` quando fornecido.
4. Usar `local_path.name` quando `remote_name` for `None`.
5. Fazer upload simples usando `PUT /content`.
6. Retornar `UploadResult`.

Na primeira versao, o upload simples deve ser usado apenas para arquivos pequenos. Um limite conservador sugerido e 4 MB.

Para arquivos maiores, o Core deve levantar erro claro:

```text
LargeFileUploadNotSupportedError
```

Upload grande com `createUploadSession` entra depois.

## Conflict Behavior

Valores aceitos inicialmente:

- `replace`: sobrescreve quando o arquivo ja existe.
- `rename`: cria novo nome quando ha conflito.
- `fail`: falha quando ha conflito.

O Core deve validar esse valor antes de chamar o Graph.

Para criar pasta, o valor deve ser enviado via:

```json
{
  "@microsoft.graph.conflictBehavior": "rename"
}
```

Para upload por caminho, o comportamento deve ser aplicado na URL quando suportado pelo Graph.

## Rotas Graph Encapsuladas

O Core deve esconder estas rotas:

```http
GET /sites/{hostname}:/{site-path}
GET /sites/{site-id}/drive
GET /drives/{drive-id}/root
GET /drives/{drive-id}/items/{folder-id}/children
POST /drives/{drive-id}/items/{folder-id}/children
PUT /drives/{drive-id}/items/{folder-id}:/{filename}:/content
PUT /drives/{drive-id}/items/{file-id}/content
```

Chamadores do Core nao devem montar URLs Graph manualmente.

## Erros

Excecoes planejadas:

```python
class MSCronosError(Exception): ...
class GraphConfigurationError(MSCronosError): ...
class SharePointUrlError(MSCronosError): ...
class GraphRequestError(MSCronosError): ...
class DriveItemNotFoundError(MSCronosError): ...
class NotAFolderError(MSCronosError): ...
class LocalFileError(MSCronosError): ...
class UnsupportedConflictBehaviorError(MSCronosError): ...
class LargeFileUploadNotSupportedError(MSCronosError): ...
```

O Core deve capturar erros conhecidos do Microsoft Graph e relancar uma excecao do projeto com contexto util.

Contexto util:

- operacao;
- site, drive ou item envolvido;
- codigo do Graph quando existir;
- mensagem original quando existir.

## Ciclo de Vida

O cliente deve ser assincrono.

Uso esperado:

```python
client = SharePointDriveClient(credentials)

try:
    site = await client.resolve_site("https://tenant.sharepoint.com/sites/RHConecta")
    drive = await client.get_default_drive(site.id)
finally:
    await client.close()
```

Opcionalmente, o Core pode implementar context manager assincrono:

```python
async with SharePointDriveClient(credentials) as client:
    ...
```

## Fluxo Principal

Fluxo esperado para enviar arquivo para uma pasta conhecida:

```text
sharepoint_url
  -> resolve_site
  -> get_default_drive
  -> get_drive_root
  -> find_child_folder("Curriculos")
  -> upload_file
```

Exemplo desejado:

```python
site = await sp.resolve_site("https://tenant.sharepoint.com/sites/RHConecta")
drive = await sp.get_default_drive(site.id)
root = await sp.get_drive_root(drive.id)
folder = await sp.find_child_folder(drive.id, root.id, "Curriculos")

if folder is None:
    folder = await sp.create_folder(drive.id, root.id, "Curriculos")

result = await sp.upload_file(
    drive_id=drive.id,
    parent_id=folder.id,
    local_path="C:/arquivos/curriculo.pdf",
    conflict_behavior="rename",
)
```

## Regras de Implementacao

- O Core deve ser pequeno.
- O Core deve ter funcoes coesas e nomes explicitos.
- O Core nao deve imprimir no terminal.
- O Core nao deve ler `.env`.
- O Core nao deve chamar `load_dotenv`.
- O Core nao deve depender de `rich`.
- O Core deve receber `GraphServiceClient` opcionalmente para facilitar testes.
- O Core deve evitar expor modelos gerados do SDK.
- O Core deve fechar credenciais e sessoes corretamente.

## Testabilidade

O Core deve permitir testes unitarios sem acesso real ao Microsoft Graph.

Estrategias:

- aceitar `GraphServiceClient` injetado;
- isolar conversao de modelos do SDK para dataclasses;
- isolar montagem de URL de upload em funcao pequena;
- validar caminhos locais com funcoes testaveis;
- criar testes unitarios para erros e validacoes.

Casos minimos de teste:

- URL SharePoint valida gera URL Graph correta;
- URL SharePoint invalida levanta `SharePointUrlError`;
- `DriveItem` de pasta vira `DriveItemRef(is_folder=True)`;
- `DriveItem` de arquivo vira `DriveItemRef(is_file=True)`;
- `conflict_behavior` invalido levanta erro;
- arquivo inexistente levanta `LocalFileError`;
- arquivo grande levanta `LargeFileUploadNotSupportedError`;
- nome remoto customizado e aplicado no upload.

## Criterios de Aceite da Primeira Versao

A primeira versao do Core sera considerada pronta quando:

1. Existir pacote `ms_cronos`.
2. Existir `SharePointDriveClient`.
3. O Core resolver site por URL.
4. O Core buscar drive padrao.
5. O Core listar filhos de uma pasta.
6. O Core criar pasta.
7. O Core enviar arquivo pequeno.
8. O Core retornar dataclasses proprias.
9. O Core nao imprimir no terminal.
10. O laboratorio `main.py` usar o Core em vez de chamar o SDK diretamente.

## Evolucao Posterior

Depois do Core basico:

1. Implementar upload grande com `createUploadSession`.
2. Criar camada de jobs.
3. Criar API simples para cadastrar jobs.
4. Criar daemon Windows.
5. Adicionar logs estruturados.
6. Adicionar retry/backoff para falhas temporarias.
7. Adicionar empacotamento para instalacao no Windows.
