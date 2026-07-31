"""Superficie publica do pacote ``core``.

Este modulo reexporta os simbolos que formam a API interna do Core para que as
demais camadas possam fazer imports simples do tipo ``from core import X``.

Regra pratica:
    Importe daqui quando estiver consumindo o Core. Importe dos submodulos
    apenas ao trabalhar dentro da propria implementacao do Core.
"""

from core.builders import build_folder_drive_item as build_folder_drive_item

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
    TreeDirectoryCreationError,
    TreeFileUploadError,
    TreeUploadError,
    UploadChunkError,
    UploadError,
    UploadSessionCreationError,
)
from core.filesystem import LocalFileSystemScanner
from core.filesystem_staging import StagingTreeBuilder

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
    RootUploadMode,
    SharePointItem,
    SharePointItemCollection,
    SharePointSite,
    SharePointSiteCollection,
    StagingFile,
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

from core.sharepoint import SharePointService

from core.urls import (
    build_create_content_url,
    build_drive_create_content_url,
    build_graph_site_url,
    validate_graph_url,
)
from core.utils import rename_with_uuid
