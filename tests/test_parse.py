"""Testes dos adaptadores entre o SDK Graph e os modelos publicos."""

from __future__ import annotations

import unittest

from msgraph.generated.models.drive import Drive
from msgraph.generated.models.drive_item import DriveItem
from msgraph.generated.models.file import File
from msgraph.generated.models.folder import Folder
from msgraph.generated.models.o_data_errors.main_error import MainError
from msgraph.generated.models.o_data_errors.o_data_error import ODataError
from msgraph.generated.models.site import Site

from core.errors import (
    FolderNotFoundError,
    GraphPermissionError,
    GraphResponseError,
)
from core.parse import parse_drive, parse_drive_item, parse_o_data_error, parse_site


class ParserTests(unittest.TestCase):
    def test_parses_site_drive_folder_and_file(self) -> None:
        site = parse_site(Site(id="site-id", name="Finance"))
        library = parse_drive(Drive(id="drive-id", name="Documents"))
        folder = parse_drive_item(
            DriveItem(id="folder-id", name="Reports", folder=Folder())
        )
        file = parse_drive_item(
            DriveItem(id="file-id", name="report.csv", file=File(), size=12)
        )

        self.assertEqual(site.id, "site-id")
        self.assertEqual(library.id, "drive-id")
        self.assertTrue(folder.is_folder)
        self.assertFalse(folder.is_file)
        self.assertTrue(file.is_file)
        self.assertEqual(file.size, 12)

    def test_requires_sdk_identifiers(self) -> None:
        with self.assertRaises(GraphResponseError):
            parse_site(Site(name="missing-id"))
        with self.assertRaises(GraphResponseError):
            parse_drive(Drive(name="missing-id"))
        with self.assertRaises(GraphResponseError):
            parse_drive_item(DriveItem(name="missing-id"))

    def test_maps_odata_code_and_operation_to_semantic_error(self) -> None:
        denied = ODataError(error=MainError(code="accessDenied", message="denied"))
        missing_folder = ODataError(
            error=MainError(code="itemNotFound", message="missing")
        )

        self.assertIsInstance(parse_o_data_error(denied), GraphPermissionError)
        self.assertIsInstance(
            parse_o_data_error(missing_folder, operation="find_folder_by_name"),
            FolderNotFoundError,
        )


if __name__ == "__main__":
    unittest.main()
