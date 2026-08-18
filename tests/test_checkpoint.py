"""Testes da persistencia de checkpoints de upload."""

from __future__ import annotations

import json
import tempfile
import unittest
from pathlib import Path, PurePosixPath

from core.checkpoint import TreeUploadCheckpointStore
from core.errors import CheckpointFormatError
from core.models import FileUploadResult, SharePointItem, TreeUploadResult


class TreeUploadCheckpointStoreTests(unittest.TestCase):
    def test_round_trip_preserves_binding_and_progress(self) -> None:
        with tempfile.TemporaryDirectory() as temporary_directory:
            checkpoint_path = Path(temporary_directory) / "upload.json"
            source_path = Path(temporary_directory) / "source.txt"
            source_path.write_text("conteudo", encoding="utf-8")
            remote_folder = SharePointItem(
                id="folder-id",
                name="docs",
                is_folder=True,
            )
            remote_file = SharePointItem(
                id="file-id",
                name="source.txt",
                is_file=True,
                size=8,
            )
            result = TreeUploadResult(
                source_root=Path(temporary_directory),
                library_id="drive-id",
                parent_item_id="root-id",
                target_root=PurePosixPath("backup"),
                staging_fingerprint="abc123",
                remote_directories={PurePosixPath("."): remote_folder},
                uploaded_files={
                    PurePosixPath("."): [
                        FileUploadResult(
                            item=remote_file,
                            source_path=source_path,
                            remote_name="source.txt",
                            conflict_behavior="replace",
                        )
                    ]
                },
                completed_levels={PurePosixPath(".")},
            )

            store = TreeUploadCheckpointStore(checkpoint_path)
            store.save(result)
            loaded = store.load()

            self.assertEqual(loaded, result)
            self.assertTrue(store.exists)

    def test_rejects_unknown_schema(self) -> None:
        with tempfile.TemporaryDirectory() as temporary_directory:
            checkpoint_path = Path(temporary_directory) / "upload.json"
            checkpoint_path.write_text(
                json.dumps({"schema_version": 999}),
                encoding="utf-8",
            )

            with self.assertRaises(CheckpointFormatError):
                TreeUploadCheckpointStore(checkpoint_path).load()


if __name__ == "__main__":
    unittest.main()
