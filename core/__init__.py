"""Superficie publica do pacote ``core``.

Este modulo reexporta os simbolos que formam a API interna do Core para que as
demais camadas possam fazer imports simples do tipo ``from core import X``.

Regra pratica:
    Importe daqui quando estiver consumindo o Core. Importe dos submodulos
    apenas ao trabalhar dentro da propria implementacao do Core.
"""

# Modelos e colecoes formam o contrato de dados mais basico do Core.
# Builders montam staging e bodies do SDK sem expor essa montagem ao servico.
from core.builders import build_folder_drive_item as build_folder_drive_item
from core.builders import build_upload_content as build_upload_content

# Erros publicos do Core, reexportados para que outras camadas nao precisem
# conhecer a estrutura interna de modulos.
from core.errors import (
    DefaultDriveNotFoundError,
    DriveItemNotFoundError,
    DriveNotFoundError,
    FailedWhenCreateDriveItemError,
    FileAlreadyExistError,
    FileVeryLargeError,
    FolderNotFoundError,
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
    NotAChildError,
    NotAFileError,
    NotAFolderError,
    SharePointUrlError,
    SiteResolutionError,
    SmallFileUploadError,
    UploadChunkError,
    UploadError,
    UploadSessionCreationError,
)
from core.filesystem import LocalFileSystemScanner

# Integracao com o Microsoft Graph: criacao do client e ciclo de vida da
# credencial assincrona.
from core.graph_client import (
    GraphClientManager,
    create_graph_client,
    create_graph_client_manager,
)
from core.models import (
    Collection_,
    CollectionItem,
    ConflictBehavior,
    DirectoryLevel,
    DirectoryLevelCollection,
    DocumentLibrary,
    DocumentLibraryCollection,
    FilesystemTree,
    FileUploadResult,
    FrozenCollection,
    GraphCredentials,
    LocalFile,
    LocalFileCollection,
    LocalFolder,
    LocalFolderCollection,
    MutableCollection,
    PreparedUpload,
    PreparedUploadCollection,
    SharePointItem,
    SharePointItemCollection,
    SharePointSite,
    SharePointSiteCollection,
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

# O servico concentra as regras de navegacao, criacao e upload sobre o client.
from core.sharepoint import SharePointService

# Helpers de URL usados para validar entradas e montar rotas especiais do
# Graph.
from core.urls import (
    build_create_content_url,
    build_drive_create_content_url,
    build_graph_site_url,
    validate_graph_url,
)
from core.utils import rename_with_uuid
