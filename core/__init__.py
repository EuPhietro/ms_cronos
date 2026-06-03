"""
Superficie publica do pacote ``core``.

Este modulo reexporta os simbolos que formam a API interna do Core para que as
demais camadas possam fazer imports simples do tipo ``from core import X``.

Regra pratica:
    Importe daqui quando estiver consumindo o Core. Importe dos submodulos apenas
    quando estiver trabalhando dentro da propria implementacao do Core.
"""

# O pacote expõe primeiro os modelos publicos porque eles servem como contrato
# de dados entre o Core e as camadas superiores.
from core.models import (
    Collection_,
    DriveItemCollection,
    DriveItemRef,
    DriveRef,
    DriveRefCollection,
    FrozenCollection,
    GraphCredentials,
    LocalFile,
    StagingContentUpload,
    LocalFileCollection,
    MutableCollection,
    SiteRef,
    SiteRefCollection,
    UploadResult,
)

# Este bloco concentra a porta de entrada da integracao com o Microsoft Graph:
# criacao do client, gerenciamento da credencial e ciclo de vida da conexao.
from core.graph_client import (
    GraphClientManager,
    create_graph_client,
    create_graph_client_manager,
)

# Aqui ficam todos os erros semanticos do projeto. Reexportar esse conjunto no
# pacote evita que o chamador precise conhecer a estrutura interna de modulos
# para capturar falhas do Core de forma consistente.
from core.errors import (
    DefaultDriveNotFoundError,
    DriveItemNotFoundError,
    DriveNotFoundError,
    FailedWhenCreateDriveItemError,
    FolderNotFoundError,
    FileVeryLargeError,
    GraphAuthenticationError,
    GraphConfigurationError,
    GraphPermissionError,
    GraphRequestError,
    GraphResourceConflictError,
    GraphResponseError,
    InvalidConflictBehaviorError,
    InvalidRemoteNameError,
    LargeFileUploadError,
    LargeFileUploadNotSupportedError,
    LocalFileNotReadableError,
    LocalPathError,
    LocalPathIsDirectoryError,
    LocalPathNotFoundError,
    NotAFileError,
    NotAFolderError,
    SharePointUrlError,
    SiteResolutionError,
    SmallFileUploadError,
    UploadChunkError,
    UploadError,
    UploadSessionCreationError,
    NotAChildError,
    FileAlreadyExistError
)

from core.parse import (
    adapt_site,
    parse_drive,
    parse_drive_collection_response,
    parse_drive_item,
    parse_drive_item_collection_response,
    parse_local_file,
    parse_o_data_error,
    parse_site,
    parse_site_collection_response,
)

# Helpers de URL usados para validar entradas do SharePoint e montar as rotas
# especiais esperadas pelo Microsoft Graph.
from core.urls import (
    build_create_content_url,
    build_drive_create_content_url,
    build_graph_site_url,
    validate_graph_url,
)

# Builders montam bodies do SDK sem expor essa montagem para o servico.
from core.builders import build_folder_drive_item, build_upload_content

# O servico de SharePoint orquestra as operacoes de leitura e escrita em cima do
# cliente autenticado do Graph.
from core.sharepoint import SharePointService

from core.utils import rename_with_uuid
