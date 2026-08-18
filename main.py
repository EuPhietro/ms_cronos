"""Exemplo executavel do upload de uma arvore local."""

import asyncio
import os
from pathlib import Path, PurePosixPath

from dotenv import load_dotenv
from rich.traceback import install

from core import (
    DriveNotFoundError,
    GraphClientManager,
    GraphCredentials,
    LocalFileSystemScanner,
    SharePointService,
    StagingTreeBuilder,
    TreeUploadProgress,
)

load_dotenv()
install(show_locals=False)


async def main() -> None:
    client_id = os.getenv("CLIENT_ID", "")
    client_secret = os.getenv("CLIENT_SECRET", "")
    client_tenant = os.getenv("CLIENT_TENANT", "")
    upload_source = os.getenv("UPLOAD_SOURCE", "")
    site_url = os.getenv("SHAREPOINT_SITE_URL", "")
    library_name = os.getenv("SHAREPOINT_LIBRARY", "")
    target_root = os.getenv("SHAREPOINT_TARGET_ROOT", ".")
    checkpoint_path = os.getenv("UPLOAD_CHECKPOINT") or None

    required = {
        "UPLOAD_SOURCE": upload_source,
        "SHAREPOINT_SITE_URL": site_url,
        "SHAREPOINT_LIBRARY": library_name,
    }
    missing = [name for name, value in required.items() if not value.strip()]
    if missing:
        raise ValueError(
            f"Defina as variaveis obrigatorias antes da execucao: {', '.join(missing)}."
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
            target_root=PurePosixPath(target_root),
        )

        site = await sharepoint.resolve_site(site_url)

        drive = await sharepoint.find_drive_by_name(site=site, name=library_name)

        if drive is None:
            raise DriveNotFoundError(
                f"A biblioteca de documentos '{library_name}' nao foi encontrada."
            )
        remote_root = await sharepoint.get_drive_root(drive)

        def show_progress(progress: TreeUploadProgress) -> None:
            print(
                f"[{progress.phase}] arquivos "
                f"{progress.completed_files}/{progress.total_files}; niveis "
                f"{progress.completed_levels}/{progress.total_levels}"
            )

        result = await sharepoint.upload_tree(
            parent=remote_root,
            library=drive,
            staging_tree=staging_tree,
            conflict_behavior="replace",
            checkpoint_path=checkpoint_path,
            progress_callback=show_progress,
        )
        print(f"Upload concluido: {result.total_uploaded_files} arquivo(s).")


if __name__ == "__main__":
    asyncio.run(main())
