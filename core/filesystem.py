"""Varredura semantica de diretorios do sistema de arquivos local.

O modulo converte uma arvore local em `FilesystemTree`, preservando cada
diretorio como um `DirectoryLevel`. Ele nao le o conteudo binario dos arquivos,
nao conhece destinos do SharePoint e nao executa uploads.

Exemplo:
    scanner = LocalFileSystemScanner()
    tree = scanner.scan("/tmp/documentos", sort_entries=True)
    print(tree.total_files)
"""

from __future__ import annotations

from collections.abc import Sequence
from pathlib import Path
from typing import Literal

from core.models import (
    DirectoryLevel,
    DirectoryLevelCollection,
    FilesystemTree,
    LocalFile,
    LocalFileCollection,
    LocalFolder,
    LocalFolderCollection,
    RootFolder,
)


class LocalFileSystemScanner:
    """Inspeciona um diretorio local e constroi um snapshot plano.

    A classe nao mantem estado entre chamadas. Cada execucao valida a raiz,
    percorre seus niveis em ordem top-down e materializa os models locais do
    Core.
    """

    def scan(
        self,
        root: Path | str,
        allow_empty: Literal["allow", "deny"] = "deny",
        sort_entries: bool = False,
    ) -> FilesystemTree:
        """Constroi um snapshot a partir de um diretorio local.

        Args:
            root: Diretorio usado como raiz da varredura.
            allow_empty: Politica repassada aos models de arquivos vazios.
            sort_entries: Ordena os itens de cada nivel para produzir um
                snapshot deterministico.

        Returns:
            Arvore plana contendo a raiz e todos os niveis encontrados.

        Raises:
            ValueError: Quando a raiz e vazia ou nao existe.
            NotADirectoryError: Quando a raiz nao representa um diretorio.
            OSError: Quando algum nivel nao pode ser lido.
        """
        if isinstance(root, str) and not root.strip():
            raise ValueError("O parâmetro path não pode ser uma str vazia")

        root = Path(root).resolve()

        if not root.exists():
            raise ValueError("O caminho não aponta para um caminho válido")

        if not root.is_dir():
            raise NotADirectoryError("O caminho não aponta para um diretório")

        root_folder = RootFolder(path=root, name=root.name)

        directory_levels: list[DirectoryLevel] = []
        for dirpath, dirnames, files in root_folder.path.walk(
            on_error=self._raise_walk_error,
            top_down=True,
        ):
            # Ordenar `dirnames` in-place tambem controla a ordem em que o
            # `Path.walk()` visitara os proximos niveis.
            if sort_entries:
                dirnames.sort()
                files.sort()

            file_collection = self._build_file_collection(
                current_path=dirpath, files=files, allow_empty=allow_empty
            )
            folder_collection = self._build_folder_collection(
                current_path=dirpath, folders=dirnames
            )

            directory_levels.append(
                self._build_directory_level(
                    current_path=dirpath,
                    files=file_collection,
                    folders=folder_collection,
                )
            )

        directory_level_collection = self._build_directory_level_collection(
            directory_levels
        )

        return FilesystemTree(root=root_folder, levels=directory_level_collection)

    def _build_directory_level_collection(
        self, directory_collection: Sequence[DirectoryLevel]
    ) -> DirectoryLevelCollection:
        """Materializa os niveis encontrados como colecao imutavel."""
        return DirectoryLevelCollection.from_collection(directory_collection)

    def _build_directory_level(
        self,
        current_path: Path,
        files: LocalFileCollection,
        folders: LocalFolderCollection,
    ) -> DirectoryLevel:
        """Combina o conteudo imediato de um diretorio em um nivel."""
        return DirectoryLevel(path=current_path, files=files, folders=folders)

    def _build_file_collection(
        self,
        files: Sequence[str],
        current_path: Path,
        allow_empty: Literal["allow", "deny"] = "deny",
    ) -> LocalFileCollection:
        """Converte os nomes de arquivos de um nivel em models locais."""
        return LocalFileCollection.from_collection(
            [
                self._build_file(
                    file, current_path=current_path, allow_empty=allow_empty
                )
                for file in files
            ]
        )

    def _build_folder_collection(
        self,
        folders: Sequence[str],
        current_path: Path,
    ) -> LocalFolderCollection:
        """Converte os nomes de subdiretorios em models locais."""
        return LocalFolderCollection.from_collection(
            [
                self._build_folder(folder, current_path=current_path)
                for folder in folders
            ]
        )

    def _build_file(
        self,
        filename: str,
        current_path: Path,
        allow_empty: Literal["allow", "deny"] = "deny",
    ) -> LocalFile:
        """Constroi a referencia de um arquivo pertencente ao nivel atual."""
        file_path = current_path / filename
        return LocalFile(
            path=file_path,
            name=file_path.name,
            size=file_path.stat().st_size,
            extension=file_path.suffix,
            allow_empty=allow_empty,
        )

    def _build_folder(self, dirname: str, current_path: Path) -> LocalFolder:
        """Constroi a referencia de um subdiretorio do nivel atual."""
        folder_path = current_path / dirname
        return LocalFolder(path=folder_path, name=folder_path.name)

    def _raise_walk_error(self, error: OSError) -> None:
        """Interrompe a varredura preservando o erro original do filesystem."""
        raise error
