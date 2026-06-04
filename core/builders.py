"""Builders dos contratos de envio usados pelo Core.

Este modulo nao executa chamadas remotas. Ele apenas:
- normaliza nomes e estrategias de conflito;
- monta bodies JSON do SDK quando o Graph espera models estruturados;
- prepara o staging interno usado pelo fluxo de upload pequeno.
"""

from typing import Literal, Optional

from msgraph.generated.models.drive_item import DriveItem
from msgraph.generated.models.folder import Folder

from core.errors import InvalidConflictBehaviorError, InvalidRemoteNameError
from core.models import LocalFile, StagingContentUpload
from core.urls import build_create_content_url


ConflictBehavior = Literal["fail", "rename", "replace"]
ALLOWED_CONFLICT_BEHAVIORS: set[str] = {"fail", "rename", "replace"}

def _normalize_remote_name(name: str) -> str:
    """Normaliza e valida um nome remoto simples de arquivo ou pasta."""
    normalized_name = name.strip()
    if not normalized_name:
        raise InvalidRemoteNameError("O nome remoto nao pode ser vazio.")
    if "/" in normalized_name or "\\" in normalized_name:
        raise InvalidRemoteNameError(
            f"O nome remoto deve ser um nome simples, nao um caminho: {name}"
        )
    return normalized_name


def _normalize_conflict_behavior(
    conflict_behavior: ConflictBehavior | str,
) -> ConflictBehavior:
    """Normaliza e valida a estrategia de conflito usada pelo Core."""
    normalized_conflict_behavior = conflict_behavior.strip().casefold()
    if normalized_conflict_behavior not in ALLOWED_CONFLICT_BEHAVIORS:
        raise InvalidConflictBehaviorError(
            "Conflict behavior invalido. Use um destes valores: fail, rename ou replace."
        )
    return normalized_conflict_behavior  # type: ignore[return-value]


def build_folder_drive_item(
    name: str,
    conflict_behavior: ConflictBehavior | str = "fail",
) -> DriveItem:
    """Monta o body `DriveItem` usado pelo Graph para criar uma pasta.

    O Graph espera um `DriveItem` com `name`, a facet `folder` e o valor especial
    `@microsoft.graph.conflictBehavior` em `additional_data`.
    """
    # Pastas remotas sao criadas por body JSON, entao o builder entrega o model
    # do SDK pronto para `children.post(...)`.
    folder_name = _normalize_remote_name(name)
    normalized_conflict_behavior = _normalize_conflict_behavior(conflict_behavior)

    body = DriveItem(name=folder_name, folder=Folder())
    body.additional_data["@microsoft.graph.conflictBehavior"] = normalized_conflict_behavior
    return body

def build_upload_content(
    file: LocalFile,
    remote_name: Optional[str],
    conflict_behavior: ConflictBehavior | str = "fail",
) -> StagingContentUpload:
    """Monta o staging semantico usado pelo upload pequeno do Core.

    O staging nao guarda a URL Graph completa. Ele armazena apenas o fragmento
    de caminho que identifica o recurso de criacao por nome, por exemplo
    `:/curriculo.pdf:/content`.
    """
    # O nome remoto pode ser sobrescrito pelo chamador ou reaproveitar o nome do
    # arquivo local quando o upload mantiver a identidade original.
    remote_file_name = _normalize_remote_name(remote_name or file.name)
    normalized_conflict_behavior = _normalize_conflict_behavior(conflict_behavior)

    return StagingContentUpload(
        file=file,
        target_path=build_create_content_url(remote_file_name),
        conflict_behavior=normalized_conflict_behavior,
    )
