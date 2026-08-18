"""Integracao Graph somente leitura, desativada por padrao.

Defina ``MS_CRONOS_RUN_INTEGRATION=1`` e as variaveis documentadas no README
para executar este teste contra um tenant controlado.
"""

from __future__ import annotations

import os
import unittest

from core import GraphClientManager, GraphCredentials, SharePointService


@unittest.skipUnless(
    os.getenv("MS_CRONOS_RUN_INTEGRATION") == "1",
    "integracao Graph desativada",
)
class ReadOnlyGraphIntegrationTests(unittest.IsolatedAsyncioTestCase):
    async def test_resolves_site_library_and_root(self) -> None:
        required_variables = (
            "CLIENT_ID",
            "CLIENT_SECRET",
            "CLIENT_TENANT",
            "MS_CRONOS_TEST_SITE_URL",
        )
        missing = [name for name in required_variables if not os.getenv(name)]
        if missing:
            self.skipTest(f"variaveis ausentes: {', '.join(missing)}")

        credentials = GraphCredentials(
            client_id=os.environ["CLIENT_ID"],
            client_secret=os.environ["CLIENT_SECRET"],
            tenant_id=os.environ["CLIENT_TENANT"],
        )
        async with GraphClientManager(credentials) as manager:
            service = SharePointService(manager)
            site = await service.resolve_site(os.environ["MS_CRONOS_TEST_SITE_URL"])
            library_name = os.getenv("MS_CRONOS_TEST_LIBRARY")
            if library_name:
                library = await service.find_drive_by_name(site, library_name)
                self.assertIsNotNone(library)
                assert library is not None
            else:
                library = await service.get_default_drive(site)
            root = await service.get_drive_root(library)

        self.assertTrue(site.id)
        self.assertTrue(library.id)
        self.assertTrue(root.id)
        self.assertTrue(root.is_folder)


if __name__ == "__main__":
    unittest.main()
