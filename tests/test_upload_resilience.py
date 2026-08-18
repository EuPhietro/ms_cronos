"""Testes dos contratos de resiliencia e upload de arvores."""

from __future__ import annotations

import asyncio
import tempfile
import unittest
from pathlib import Path, PurePosixPath
from types import SimpleNamespace
from unittest.mock import AsyncMock, patch

import httpx
from msgraph.generated.models.o_data_errors.o_data_error import ODataError

from core.checkpoint import TreeUploadCheckpointStore
from core.errors import (
    CheckpointMismatchError,
    FileAlreadyExistError,
    InvalidRemoteNameError,
    TreeFileUploadError,
    TreeUploadCancelledError,
)
from core.models import (
    DocumentLibrary,
    FileUploadResult,
    LocalFile,
    SharePointItem,
    SharePointItemCollection,
    TreeUploadResult,
)
from core.sharepoint import SharePointService
from core.urls import validate_remote_name, validate_remote_path


class GraphRetryTests(unittest.IsolatedAsyncioTestCase):
    """Verifica retries sem realizar requisicoes externas."""

    async def test_retries_throttled_operation_using_retry_after(self) -> None:
        service = SharePointService(SimpleNamespace())
        operation = AsyncMock(
            side_effect=[
                ODataError(
                    message="throttled",
                    response_status_code=429,
                    response_headers={"Retry-After": "2"},
                ),
                "ok",
            ]
        )

        with patch("core.sharepoint.asyncio.sleep", new_callable=AsyncMock) as sleep:
            result = await service._execute_graph_operation(
                operation,
                operation_name="teste de throttling",
            )

        self.assertEqual(result, "ok")
        self.assertEqual(operation.await_count, 2)
        sleep.assert_awaited_once_with(2.0)

    async def test_retries_transport_failure_with_exponential_delay(self) -> None:
        service = SharePointService(SimpleNamespace())
        request = httpx.Request("GET", "https://graph.microsoft.com/v1.0/")
        operation = AsyncMock(
            side_effect=[httpx.ReadError("connection reset", request=request), "ok"]
        )

        with patch("core.sharepoint.asyncio.sleep", new_callable=AsyncMock) as sleep:
            result = await service._execute_graph_operation(
                operation,
                operation_name="teste de transporte",
            )

        self.assertEqual(result, "ok")
        sleep.assert_awaited_once_with(1.0)


class RemoteValidationTests(unittest.TestCase):
    """Cobre nomes e fragmentos rejeitados antes da chamada ao Graph."""

    def test_rejects_invalid_remote_names(self) -> None:
        for name in (
            "",
            " arquivo.txt",
            "arquivo?.txt",
            "CON",
            "con.txt",
            "pasta.",
        ):
            with self.subTest(name=name):
                with self.assertRaises(InvalidRemoteNameError):
                    validate_remote_name(name)

    def test_rejects_absolute_and_overlong_remote_paths(self) -> None:
        with self.assertRaises(InvalidRemoteNameError):
            validate_remote_path(PurePosixPath("/absoluto"))

        overlong_path = PurePosixPath("a" * 201) / ("b" * 200)
        with self.assertRaises(InvalidRemoteNameError):
            validate_remote_path(overlong_path)


class TreeUploadTests(unittest.IsolatedAsyncioTestCase):
    """Verifica cache top-down e checkpoint sem acessar o SharePoint."""

    async def test_lists_each_level_once_and_reuses_completed_checkpoint(self) -> None:
        service = SharePointService(SimpleNamespace())
        library = DocumentLibrary(id="drive-id", name="Documents")
        remote_root = SharePointItem(id="root-id", name="root", is_folder=True)
        remote_docs = SharePointItem(id="docs-id", name="docs", is_folder=True)

        with tempfile.TemporaryDirectory() as temporary_directory:
            root_path = Path(temporary_directory)
            root_file = root_path / "root.txt"
            root_file.write_bytes(b"root")
            docs_path = root_path / "docs"
            docs_path.mkdir()
            docs_file = docs_path / "child.txt"
            docs_file.write_bytes(b"child")

            root_level = SimpleNamespace(
                relative_path=PurePosixPath("."),
                staging_folders=(
                    SimpleNamespace(
                        relative_path=PurePosixPath("docs"),
                        remote_name="docs",
                    ),
                ),
                staging_files=(self._staging_file(root_file, "root.txt"),),
            )
            docs_level = SimpleNamespace(
                relative_path=PurePosixPath("docs"),
                staging_folders=(),
                staging_files=(self._staging_file(docs_file, "docs/child.txt"),),
            )
            staging_tree = SimpleNamespace(
                target_root=PurePosixPath("."),
                levels=(root_level, docs_level),
                source=SimpleNamespace(root=SimpleNamespace(path=root_path)),
            )

            service.ensure_remote_folder_path = AsyncMock(return_value=remote_root)
            service.list_children = AsyncMock(
                return_value=SharePointItemCollection.from_collection([])
            )
            service.create_folder = AsyncMock(return_value=remote_docs)
            service._upload_small_file = AsyncMock(
                side_effect=self._successful_small_upload
            )
            progress = []

            checkpoint = await service.upload_tree(
                remote_root,
                library,
                staging_tree,
                progress_callback=progress.append,
            )

            self.assertEqual(service.list_children.await_count, 2)
            service.create_folder.assert_awaited_once()
            self.assertEqual(service._upload_small_file.await_count, 2)
            self.assertEqual(checkpoint.total_uploaded_files, 2)
            self.assertEqual(
                [update.phase for update in progress],
                [
                    "preparing_directories",
                    "uploading_files",
                    "uploading_files",
                    "uploading_files",
                    "completed",
                ],
            )
            self.assertEqual(progress[-1].completed_files, 2)

            resumed = await service.upload_tree(
                remote_root,
                library,
                staging_tree,
                checkpoint=checkpoint,
            )

            self.assertIs(resumed, checkpoint)
            self.assertEqual(service.create_folder.await_count, 1)
            self.assertEqual(service._upload_small_file.await_count, 2)

    async def test_cancellation_exposes_bound_partial_result(self) -> None:
        service = SharePointService(SimpleNamespace())
        library = DocumentLibrary(id="drive-id")
        remote_root = SharePointItem(id="root-id", is_folder=True)
        cancel_event = asyncio.Event()
        cancel_event.set()

        with tempfile.TemporaryDirectory() as temporary_directory:
            checkpoint_path = Path(temporary_directory) / "cancelled.json"
            staging_tree = SimpleNamespace(
                target_root=PurePosixPath("backup"),
                levels=(),
                source=SimpleNamespace(
                    root=SimpleNamespace(path=Path(temporary_directory))
                ),
            )

            with self.assertRaises(TreeUploadCancelledError) as raised:
                await service.upload_tree(
                    remote_root,
                    library,
                    staging_tree,
                    cancel_event=cancel_event,
                    checkpoint_path=checkpoint_path,
                )

            self.assertTrue(checkpoint_path.is_file())

        partial_result = raised.exception.partial_result
        self.assertEqual(partial_result.library_id, "drive-id")
        self.assertEqual(partial_result.parent_item_id, "root-id")

    async def test_rejects_checkpoint_bound_to_another_library(self) -> None:
        service = SharePointService(SimpleNamespace())
        library = DocumentLibrary(id="drive-id")
        remote_root = SharePointItem(id="root-id", is_folder=True)

        with tempfile.TemporaryDirectory() as temporary_directory:
            staging_tree = SimpleNamespace(
                target_root=PurePosixPath("."),
                levels=(),
                source=SimpleNamespace(
                    root=SimpleNamespace(path=Path(temporary_directory))
                ),
            )
            checkpoint = TreeUploadResult(library_id="another-drive")

            with self.assertRaises(CheckpointMismatchError):
                await service.upload_tree(
                    remote_root,
                    library,
                    staging_tree,
                    checkpoint=checkpoint,
                )

    async def test_loads_completed_checkpoint_from_path(self) -> None:
        service = SharePointService(SimpleNamespace())
        library = DocumentLibrary(id="drive-id")
        remote_root = SharePointItem(id="root-id", is_folder=True)

        with tempfile.TemporaryDirectory() as temporary_directory:
            root_path = Path(temporary_directory)
            checkpoint_path = root_path / "upload.json"
            staging_tree = SimpleNamespace(
                target_root=PurePosixPath("."),
                levels=(),
                source=SimpleNamespace(root=SimpleNamespace(path=root_path)),
            )
            checkpoint = TreeUploadResult(
                source_root=root_path.resolve(),
                library_id=library.id,
                parent_item_id=remote_root.id,
                target_root=PurePosixPath("."),
                staging_fingerprint=service._staging_fingerprint(staging_tree),
                remote_directories={PurePosixPath("."): remote_root},
            )
            TreeUploadCheckpointStore(checkpoint_path).save(checkpoint)

            resumed = await service.upload_tree(
                remote_root,
                library,
                staging_tree,
                checkpoint_path=checkpoint_path,
            )

            self.assertEqual(resumed, checkpoint)

    async def test_exposes_partial_result_when_a_file_fails(self) -> None:
        service = SharePointService(SimpleNamespace())
        library = DocumentLibrary(id="drive-id")
        remote_root = SharePointItem(id="root-id", name="root", is_folder=True)

        with tempfile.TemporaryDirectory() as temporary_directory:
            root_path = Path(temporary_directory)
            first_path = root_path / "first.txt"
            second_path = root_path / "second.txt"
            first_path.write_bytes(b"first")
            second_path.write_bytes(b"second")
            level = SimpleNamespace(
                relative_path=PurePosixPath("."),
                staging_folders=(),
                staging_files=(
                    self._staging_file(first_path, "first.txt"),
                    self._staging_file(second_path, "second.txt"),
                ),
            )
            staging_tree = SimpleNamespace(
                target_root=PurePosixPath("."),
                levels=(level,),
                source=SimpleNamespace(root=SimpleNamespace(path=root_path)),
            )

            service.ensure_remote_folder_path = AsyncMock(return_value=remote_root)
            service.list_children = AsyncMock(
                return_value=SharePointItemCollection.from_collection([])
            )
            service._upload_small_file = AsyncMock(
                side_effect=[
                    await self._successful_small_upload(
                        local_file=level.staging_files[0].source,
                        remote_name="first.txt",
                        conflict_behavior="fail",
                    ),
                    RuntimeError("falha simulada"),
                ]
            )

            with self.assertRaises(TreeFileUploadError) as raised:
                await service.upload_tree(remote_root, library, staging_tree)

            partial_result = raised.exception.partial_result
            self.assertIsNotNone(partial_result)
            self.assertEqual(partial_result.total_uploaded_files, 1)
            self.assertNotIn(PurePosixPath("."), partial_result.completed_levels)

    async def test_small_upload_fail_policy_rejects_existing_name(self) -> None:
        service = SharePointService(SimpleNamespace())
        library = DocumentLibrary(id="drive-id")
        remote_root = SharePointItem(id="root-id", is_folder=True)
        existing_file = SharePointItem(
            id="existing-id",
            name="report.txt",
            is_file=True,
        )

        with tempfile.TemporaryDirectory() as temporary_directory:
            source_path = Path(temporary_directory) / "report.txt"
            source_path.write_text("report", encoding="utf-8")

            with self.assertRaises(FileAlreadyExistError):
                await service._upload_small_file(
                    library,
                    remote_root,
                    LocalFile.from_uri(source_path),
                    conflict_behavior="fail",
                    existing_file=existing_file,
                    remote_state_known=True,
                )

    @staticmethod
    def _staging_file(path: Path, relative_path: str) -> SimpleNamespace:
        return SimpleNamespace(
            source=SimpleNamespace(path=path, size=path.stat().st_size, name=path.name),
            relative_path=PurePosixPath(relative_path),
            remote_name=path.name,
            conflict_behavior="fail",
        )

    @staticmethod
    async def _successful_small_upload(
        *,
        local_file: SimpleNamespace,
        remote_name: str,
        conflict_behavior: str,
        **_: object,
    ) -> FileUploadResult:
        return FileUploadResult(
            item=SharePointItem(
                id=f"remote-{remote_name}",
                name=remote_name,
                is_file=True,
            ),
            source_path=local_file.path,
            remote_name=remote_name,
            conflict_behavior=conflict_behavior,
        )


if __name__ == "__main__":
    unittest.main()
