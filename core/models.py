from pathlib import Path
from dataclasses import dataclass



dataclass(frozen=True)
class GraphCredentials:
    '''
    dataclasse que define as credenciais de acesso
    '''
    client_id: str
    secrets_token: str
    tenant_id: str
    
dataclass(frozen=True)
class SiteRef:
    id: str
    name: str | None
    display_name: str | None
    web_url: str | None
    
    
dataclass(frozen=True)
class DriveItemRef:
    id: str
    name: str | None
    web_url: str | None
    is_folder: bool
    is_file: bool
    size: int | None


dataclass(frozen=True)
class UploadResult:
    item: DriveItemRef
    source_path: Path
    remote_name: str
    conflict_behavior: str
    