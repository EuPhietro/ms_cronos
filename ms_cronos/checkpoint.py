"""Persistencia JSON do progresso de uploads de arvore.

O arquivo de checkpoint contem apenas metadados e referencias semanticas. Ele
nao armazena credenciais, conteudo de arquivos nem URLs de upload session.
"""

from __future__ import annotations

import json
import os
from pathlib import Path, PurePosixPath
from tempfile import NamedTemporaryFile
from typing import Any, cast

from core.errors import CheckpointError, CheckpointFormatError
from core.models import (
    ConflictBehavior,
    FileUploadResult,
    SharePointItem,
    TreeUploadResult,
)

CHECKPOINT_SCHEMA_VERSION = 1


class TreeUploadCheckpointStore:
    """Le e grava checkpoints de forma atomica em um arquivo JSON."""

    def __init__(self, path: str | Path) -> None:
        self.path = Path(path).expanduser()

    @property
    def exists(self) -> bool:
        """Indica se o arquivo de checkpoint ja existe."""
        return self.path.is_file()

    def save(self, result: TreeUploadResult) -> None:
        """Grava o estado em arquivo temporario e o substitui atomicamente."""
        self.path.parent.mkdir(parents=True, exist_ok=True)
        payload = self._serialize(result)
        temporary_path: Path | None = None

        try:
            with NamedTemporaryFile(
                mode="w",
                encoding="utf-8",
                dir=self.path.parent,
                prefix=f".{self.path.name}.",
                suffix=".tmp",
                delete=False,
            ) as temporary_file:
                temporary_path = Path(temporary_file.name)
                json.dump(payload, temporary_file, ensure_ascii=False, indent=2)
                temporary_file.write("\n")
                temporary_file.flush()
                os.fsync(temporary_file.fileno())

            os.replace(temporary_path, self.path)
        except OSError as error:
            if temporary_path is not None:
                temporary_path.unlink(missing_ok=True)
            raise CheckpointError(
                f"Nao foi possivel gravar o checkpoint em '{self.path}'."
            ) from error

    def load(self) -> TreeUploadResult:
        """Carrega e valida um checkpoint previamente gravado."""
        try:
            with self.path.open(encoding="utf-8") as checkpoint_file:
                payload = json.load(checkpoint_file)
        except (OSError, json.JSONDecodeError) as error:
            raise CheckpointFormatError(
                f"Nao foi possivel ler um checkpoint JSON valido em '{self.path}'."
            ) from error

        try:
            return self._deserialize(payload)
        except (KeyError, TypeError, ValueError) as error:
            raise CheckpointFormatError(
                f"O checkpoint '{self.path}' nao possui o schema esperado."
            ) from error

    @staticmethod
    def _serialize(result: TreeUploadResult) -> dict[str, Any]:
        return {
            "schema_version": CHECKPOINT_SCHEMA_VERSION,
            "binding": {
                "source_root": str(result.source_root)
                if result.source_root is not None
                else None,
                "library_id": result.library_id,
                "parent_item_id": result.parent_item_id,
                "target_root": result.target_root.as_posix()
                if result.target_root is not None
                else None,
                "staging_fingerprint": result.staging_fingerprint,
            },
            "remote_directories": {
                path.as_posix(): TreeUploadCheckpointStore._serialize_item(item)
                for path, item in result.remote_directories.items()
            },
            "uploaded_files": {
                path.as_posix(): [
                    {
                        "item": TreeUploadCheckpointStore._serialize_item(upload.item),
                        "source_path": str(upload.source_path),
                        "remote_name": upload.remote_name,
                        "conflict_behavior": upload.conflict_behavior,
                    }
                    for upload in uploads
                ]
                for path, uploads in result.uploaded_files.items()
            },
            "completed_levels": sorted(
                path.as_posix() for path in result.completed_levels
            ),
        }

    @staticmethod
    def _deserialize(payload: object) -> TreeUploadResult:
        if not isinstance(payload, dict):
            raise TypeError("O checkpoint deve ser um objeto JSON.")
        if payload.get("schema_version") != CHECKPOINT_SCHEMA_VERSION:
            raise ValueError("Versao de schema de checkpoint nao suportada.")

        binding = TreeUploadCheckpointStore._require_dict(payload, "binding")
        directories = TreeUploadCheckpointStore._require_dict(
            payload, "remote_directories"
        )
        uploaded_files = TreeUploadCheckpointStore._require_dict(
            payload, "uploaded_files"
        )
        completed_levels = payload["completed_levels"]
        if not isinstance(completed_levels, list):
            raise TypeError("completed_levels deve ser uma lista.")

        source_root = binding.get("source_root")
        target_root = binding.get("target_root")
        result = TreeUploadResult(
            source_root=Path(source_root) if isinstance(source_root, str) else None,
            library_id=TreeUploadCheckpointStore._optional_string(
                binding, "library_id"
            ),
            parent_item_id=TreeUploadCheckpointStore._optional_string(
                binding, "parent_item_id"
            ),
            target_root=PurePosixPath(target_root)
            if isinstance(target_root, str)
            else None,
            staging_fingerprint=TreeUploadCheckpointStore._optional_string(
                binding, "staging_fingerprint"
            ),
            remote_directories={
                PurePosixPath(path): TreeUploadCheckpointStore._deserialize_item(item)
                for path, item in directories.items()
            },
            completed_levels={
                PurePosixPath(TreeUploadCheckpointStore._require_string(path))
                for path in completed_levels
            },
        )

        for path, uploads in uploaded_files.items():
            if not isinstance(path, str) or not isinstance(uploads, list):
                raise TypeError("uploaded_files possui uma entrada invalida.")
            result.uploaded_files[PurePosixPath(path)] = [
                TreeUploadCheckpointStore._deserialize_upload(upload)
                for upload in uploads
            ]
        return result

    @staticmethod
    def _serialize_item(item: SharePointItem) -> dict[str, Any]:
        return {
            "id": item.id,
            "name": item.name,
            "web_url": item.web_url,
            "is_folder": item.is_folder,
            "is_file": item.is_file,
            "size": item.size,
        }

    @staticmethod
    def _deserialize_item(payload: object) -> SharePointItem:
        if not isinstance(payload, dict):
            raise TypeError("O item remoto do checkpoint deve ser um objeto.")
        return SharePointItem(
            id=TreeUploadCheckpointStore._required_string(payload, "id"),
            name=TreeUploadCheckpointStore._optional_string(payload, "name"),
            web_url=TreeUploadCheckpointStore._optional_string(payload, "web_url"),
            is_folder=TreeUploadCheckpointStore._required_bool(payload, "is_folder"),
            is_file=TreeUploadCheckpointStore._required_bool(payload, "is_file"),
            size=TreeUploadCheckpointStore._optional_int(payload, "size"),
        )

    @staticmethod
    def _deserialize_upload(payload: object) -> FileUploadResult:
        if not isinstance(payload, dict):
            raise TypeError("O upload do checkpoint deve ser um objeto.")
        behavior = TreeUploadCheckpointStore._required_string(
            payload, "conflict_behavior"
        )
        if behavior not in {"fail", "rename", "replace"}:
            raise ValueError("Conflict behavior invalido no checkpoint.")
        return FileUploadResult(
            item=TreeUploadCheckpointStore._deserialize_item(payload["item"]),
            source_path=Path(
                TreeUploadCheckpointStore._required_string(payload, "source_path")
            ),
            remote_name=TreeUploadCheckpointStore._required_string(
                payload, "remote_name"
            ),
            conflict_behavior=cast(ConflictBehavior, behavior),
        )

    @staticmethod
    def _require_dict(payload: dict[str, Any], key: str) -> dict[str, Any]:
        value = payload[key]
        if not isinstance(value, dict):
            raise TypeError(f"{key} deve ser um objeto.")
        return value

    @staticmethod
    def _require_string(value: object) -> str:
        if not isinstance(value, str):
            raise TypeError("O valor deve ser uma string.")
        return value

    @staticmethod
    def _required_string(payload: dict[str, Any], key: str) -> str:
        return TreeUploadCheckpointStore._require_string(payload[key])

    @staticmethod
    def _optional_string(payload: dict[str, Any], key: str) -> str | None:
        value = payload.get(key)
        if value is not None and not isinstance(value, str):
            raise TypeError(f"{key} deve ser uma string ou null.")
        return value

    @staticmethod
    def _required_bool(payload: dict[str, Any], key: str) -> bool:
        value = payload[key]
        if not isinstance(value, bool):
            raise TypeError(f"{key} deve ser booleano.")
        return value

    @staticmethod
    def _optional_int(payload: dict[str, Any], key: str) -> int | None:
        value = payload.get(key)
        if value is not None and (
            not isinstance(value, int) or isinstance(value, bool)
        ):
            raise TypeError(f"{key} deve ser inteiro ou null.")
        return value
