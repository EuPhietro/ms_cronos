"""Testes do scanner local e da transformacao para staging."""

from __future__ import annotations

import tempfile
import unittest
from pathlib import Path, PurePosixPath

from core.errors import LocalFileNotReadableError, LocalPathNotFoundError
from core.filesystem import LocalFileSystemScanner
from core.filesystem_staging import StagingTreeBuilder
from core.models import LocalFile


class FilesystemPipelineTests(unittest.TestCase):
    def test_scans_and_stages_nested_and_empty_directories(self) -> None:
        with tempfile.TemporaryDirectory() as temporary_directory:
            root = Path(temporary_directory)
            (root / "empty").mkdir()
            nested = root / "docs" / "reports"
            nested.mkdir(parents=True)
            (root / "root.txt").write_text("root", encoding="utf-8")
            (nested / "report.csv").write_text("a,b\n1,2", encoding="utf-8")

            tree = LocalFileSystemScanner().scan(root, sort_entries=True)
            staging = StagingTreeBuilder().build_staging_tree(
                tree,
                conflict_behavior="replace",
                target_root=PurePosixPath("backup"),
            )

            self.assertEqual(tree.total_files, 2)
            self.assertEqual(tree.total_levels, 4)
            self.assertEqual(tree.total_subdirectories, 3)
            self.assertEqual(len(staging.levels), tree.total_levels)
            self.assertEqual(staging.levels[0].relative_path, PurePosixPath("."))
            self.assertEqual(staging.target_root, PurePosixPath("backup"))
            self.assertEqual(
                {
                    level.relative_path
                    for level in staging.levels
                    if not level.staging_files and not level.staging_folders
                },
                {PurePosixPath("empty")},
            )

    def test_empty_file_policy_uses_semantic_error(self) -> None:
        with tempfile.TemporaryDirectory() as temporary_directory:
            empty_file = Path(temporary_directory) / "empty.txt"
            empty_file.touch()

            with self.assertRaises(LocalFileNotReadableError):
                LocalFile.from_uri(empty_file)

            allowed = LocalFile(
                path=empty_file,
                name=empty_file.name,
                size=0,
                extension=empty_file.suffix,
                allow_empty="allow",
            )
            self.assertEqual(allowed.size, 0)

    def test_missing_file_uses_semantic_error(self) -> None:
        with tempfile.TemporaryDirectory() as temporary_directory:
            missing = Path(temporary_directory) / "missing.txt"
            with self.assertRaises(LocalPathNotFoundError):
                LocalFile.from_uri(missing)


if __name__ == "__main__":
    unittest.main()
