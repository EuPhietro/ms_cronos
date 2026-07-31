"""Execucao manual simples do scanner de diretorios locais."""

import asyncio
import os
from pathlib import Path, PurePosixPath

from dotenv import load_dotenv
from rich.traceback import install

from core import LocalFileSystemScanner, StagingTreeBuilder
from core.errors import DriveNotFoundError
from core.graph_client import GraphClientManager
from core.models import GraphCredentials
from core.sharepoint import SharePointService

load_dotenv()
install(show_locals=False)


async def main() -> None:
    client_id = os.getenv("CLIENT_ID", "")
    client_secret = os.getenv("CLIENT_SECRET", "")
    client_tenant = os.getenv("CLIENT_TENANT", "")
    upload_source = os.getenv("UPLOAD_SOURCE", "")

    if not upload_source.strip():
        raise ValueError(
            "Defina UPLOAD_SOURCE com o diretorio local que deve ser enviado."
        )

    graph_credentials = GraphCredentials(
        client_id=client_id,
        client_secret=client_secret,
        tenant_id=client_tenant,
    )

    async with GraphClientManager(credentials=graph_credentials) as client:
        # A origem explicita evita enviar `.env`, `venv` e caches do projeto
        # acidentalmente ao usar o diretorio deste script como raiz.
        root = Path(upload_source).expanduser().resolve()
        scanner = LocalFileSystemScanner()
        builder = StagingTreeBuilder()
        sharepoint = SharePointService(graph_client_manager=client)
        tree = scanner.scan(root, sort_entries=True, allow_empty="allow")
        print(tree.total_files)
        print(tree.total_levels)
        print(tree.total_size)
        print(tree.total_subdirectories)
        staging_tree = builder.build_staging_tree(
            tree=tree,
            conflict_behavior="replace",
            target_root=PurePosixPath("ms_cronos"),
        )

        site = await sharepoint.resolve_site(
            "https://plangeconcombr.sharepoint.com/sites/RHConecta"
        )

        drive = await sharepoint.find_drive_by_name(site=site, name="SESMT")

        if drive is None:
            raise DriveNotFoundError(
                "A biblioteca de documentos 'SESMT' nao foi encontrada no site."
            )
        remote_root = await sharepoint.get_drive_root(drive)

        await sharepoint.upload_tree(
            parent=remote_root,
            library=drive,
            staging_tree=staging_tree,
            conflict_behavior="replace",
        )


if __name__ == "__main__":
    asyncio.run(main())
