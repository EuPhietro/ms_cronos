"""Testes da superficie publica e do versionamento Beta 1."""

from __future__ import annotations

import tomllib
import unittest
from pathlib import Path

import core


class PublicApiTests(unittest.TestCase):
    def test_public_symbols_are_resolvable(self) -> None:
        for symbol in core.__all__:
            with self.subTest(symbol=symbol):
                self.assertTrue(hasattr(core, symbol))

    def test_runtime_and_package_versions_match(self) -> None:
        project_root = Path(__file__).resolve().parents[1]
        with (project_root / "pyproject.toml").open("rb") as pyproject_file:
            package_version = tomllib.load(pyproject_file)["project"]["version"]

        self.assertEqual(core.__version__, "0.1.0b1")
        self.assertEqual(core.__version__, package_version)

    def test_sdk_models_are_not_exported(self) -> None:
        self.assertNotIn("DriveItem", core.__all__)
        self.assertNotIn("Drive", core.__all__)
        self.assertNotIn("ODataError", core.__all__)


if __name__ == "__main__":
    unittest.main()
