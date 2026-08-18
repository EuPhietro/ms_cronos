"""Testes da paginacao encapsulada pelo SharePointService."""

from __future__ import annotations

import unittest
from types import SimpleNamespace
from unittest.mock import AsyncMock, MagicMock, patch

from msgraph.generated.models.drive import Drive
from msgraph.generated.models.drive_collection_response import DriveCollectionResponse
from msgraph.generated.models.drive_item import DriveItem
from msgraph.generated.models.drive_item_collection_response import (
    DriveItemCollectionResponse,
)
from msgraph.generated.models.folder import Folder

from core.models import DocumentLibrary, SharePointItem, SharePointSite
from core.sharepoint import SharePointService


class _FakePageIterator:
    def __init__(self, response: object, pages: list[object | None]) -> None:
        self.current_page = response
        self._pages = pages
        self.next_calls = 0

    async def next(self) -> object | None:
        self.next_calls += 1
        return self._pages.pop(0) if self._pages else None


class PaginationTests(unittest.IsolatedAsyncioTestCase):
    async def test_list_site_drives_accumulates_pages_without_syncing_iterator(
        self,
    ) -> None:
        first = DriveCollectionResponse(value=[Drive(id="drive-1", name="Documents")])
        second = DriveCollectionResponse(value=[Drive(id="drive-2", name="Archive")])
        iterator = _FakePageIterator(first, [second, None])
        service = SharePointService(self._manager_for_drives(first))

        with patch("core.sharepoint.PageIterator", return_value=iterator):
            libraries = await service.list_site_drives(SharePointSite(id="site-id"))

        self.assertEqual([library.id for library in libraries], ["drive-1", "drive-2"])
        self.assertEqual(iterator.next_calls, 2)

    async def test_list_site_drives_honors_client_side_page_limit(self) -> None:
        first = DriveCollectionResponse(value=[Drive(id="drive-1")])
        iterator = _FakePageIterator(
            first,
            [DriveCollectionResponse(value=[Drive(id="drive-2")])],
        )
        service = SharePointService(self._manager_for_drives(first))

        with patch("core.sharepoint.PageIterator", return_value=iterator):
            libraries = await service.list_site_drives(
                SharePointSite(id="site-id"),
                max_pages=1,
            )

        self.assertEqual([library.id for library in libraries], ["drive-1"])
        self.assertEqual(iterator.next_calls, 0)

    async def test_iter_children_continues_after_empty_intermediate_page(self) -> None:
        first = DriveItemCollectionResponse(
            value=[DriveItem(id="folder-1", name="one", folder=Folder())]
        )
        empty = DriveItemCollectionResponse(value=[])
        third = DriveItemCollectionResponse(
            value=[DriveItem(id="folder-2", name="two", folder=Folder())]
        )
        iterator = _FakePageIterator(first, [empty, third, None])
        service = SharePointService(self._manager_for_children(first))

        with patch("core.sharepoint.PageIterator", return_value=iterator):
            pages = [
                page
                async for page in service.iter_children(
                    DocumentLibrary(id="drive-id"),
                    SharePointItem(id="root-id", is_folder=True),
                )
            ]

        self.assertEqual(
            [[item.id for item in page] for page in pages], [["folder-1"], ["folder-2"]]
        )
        self.assertEqual(iterator.next_calls, 3)

    @staticmethod
    def _manager_for_drives(response: DriveCollectionResponse) -> SimpleNamespace:
        drives = SimpleNamespace(get=AsyncMock(return_value=response))
        sites = SimpleNamespace(
            by_site_id=MagicMock(return_value=SimpleNamespace(drives=drives))
        )
        client = SimpleNamespace(sites=sites, request_adapter=object())
        return SimpleNamespace(client=client)

    @staticmethod
    def _manager_for_children(
        response: DriveItemCollectionResponse,
    ) -> SimpleNamespace:
        children = SimpleNamespace(get=AsyncMock(return_value=response))
        item_endpoint = SimpleNamespace(children=children)
        items = SimpleNamespace(by_drive_item_id=MagicMock(return_value=item_endpoint))
        drive_endpoint = SimpleNamespace(items=items)
        drives = SimpleNamespace(by_drive_id=MagicMock(return_value=drive_endpoint))
        client = SimpleNamespace(drives=drives, request_adapter=object())
        return SimpleNamespace(client=client)


if __name__ == "__main__":
    unittest.main()
