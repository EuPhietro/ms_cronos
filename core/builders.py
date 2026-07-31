"""Builders e normalizadores dos contratos de envio usados pelo Core.

Este modulo nao executa chamadas remotas. Ele apenas:
- normaliza nomes de recursos remotos;
- monta bodies JSON do SDK quando o Graph espera models estruturados;
"""

from typing import cast

from msgraph.generated.models.drive_item import DriveItem
from msgraph.generated.models.folder import Folder

from core.errors import InvalidConflictBehaviorError, InvalidRemoteNameError
from core.models import ConflictBehavior

ALLOWED_CONFLICT_BEHAVIORS = frozenset({"fail", "rename", "replace"})


def normalize_conflict_behavior(conflict_behavior: str) -> ConflictBehavior:
    """Normaliza e valida uma politica de conflito recebida em runtime."""
    normalized_behavior = conflict_behavior.strip().casefold()
    if normalized_behavior not in ALLOWED_CONFLICT_BEHAVIORS:
        raise InvalidConflictBehaviorError(
            "Conflict behavior invalido. Use 'fail', 'rename' ou 'replace'."
        )
    return cast(ConflictBehavior, normalized_behavior)


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


def build_folder_drive_item(
    name: str,
    conflict_behavior: ConflictBehavior = "fail",
) -> DriveItem:
    """Monta o body `DriveItem` usado pelo Graph para criar uma pasta.

    O Graph espera um `DriveItem` com `name`, a facet `folder` e o valor
    especial `@microsoft.graph.conflictBehavior` em `additional_data`.
    """
    # Pastas remotas sao criadas por body JSON, entao o builder entrega o model
    # do SDK pronto para `children.post(...)`.
    folder_name = _normalize_remote_name(name)
    normalized_behavior = normalize_conflict_behavior(conflict_behavior)

    body = DriveItem(name=folder_name, folder=Folder())
    body.additional_data["@microsoft.graph.conflictBehavior"] = normalized_behavior
    return body
