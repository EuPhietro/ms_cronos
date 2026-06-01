"""Superficie publica do pacote ``core``.

Este modulo reexporta os simbolos que formam a API interna do Core para que as
demais camadas possam fazer imports simples do tipo ``from core import X``.
"""

# O pacote expõe primeiro os modelos publicos porque eles servem como contrato
# de dados entre o Core e as camadas superiores.
from core.models import (
    GraphCredentials,
    SiteRef,
    DriveRef,
    DriveItemRef,
    LocalFile,
    UploadResult,
    Collection_,
    MutableCollection,
    FrozenCollection,
    SiteRefCollection,
    DriveRefCollection,
    DriveItemCollection,
    LocalFileCollection 
    )

# Este bloco concentra a porta de entrada da integracao com o Microsoft Graph:
# criacao do client, gerenciamento da credencial e ciclo de vida da conexao.
from core.graph_client import (GraphClientManager, create_graph_client, create_graph_client_manager)

# Aqui ficam todos os erros semanticos do projeto. Reexportar esse conjunto no
# pacote evita que o chamador precise conhecer a estrutura interna de modulos
# para capturar falhas do Core de forma consistente.
from core.errors import (
    GraphConfigurationError,
    GraphAuthenticationError,
    SharePointUrlError,
    InvalidConflictBehaviorError,
    InvalidRemoteNameError,
    LocalPathError,
    LocalPathNotFoundError,
    LocalPathIsDirectoryError,
    LocalFileNotReadableError,
    SiteResolutionError,
    DriveNotFoundError,
    DefaultDriveNotFoundError,
    DriveItemNotFoundError,
    FolderNotFoundError,
    NotAFolderError,
    NotAFileError,
    GraphRequestError,
    GraphResourceConflictError,
    GraphPermissionError,
    GraphResponseError,
    UploadError,
    SmallFileUploadError,
    LargeFileUploadError,
    LargeFileUploadNotSupportedError,
    UploadChunkError,
    UploadSessionCreationError
    )

from core.parse import (
    parse_site,
    parse_drive,
    parse_drive_item,
    parse_local_file,
    parse_site_collection_response,
    parse_drive_collection_response,
    parse_drive_item_collection_response,
    parse_o_data_error,
    adapt_site
    )

# Por fim, o pacote expõe os helpers de URL usados para validar entradas do
# SharePoint e montar as rotas especiais esperadas pelo Microsoft Graph.
from core.urls import build_graph_site_url, validate_graph_url

# O servico de SharePoint orquestra as operacoes de leitura e, futuramente, as
# operacoes de escrita em cima do cliente autenticado do Graph.
from core.sharepoint import SharePointService
