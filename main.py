"""Execucao manual simples do scanner de diretorios locais."""

import asyncio
from pathlib import Path

from dotenv import load_dotenv
from rich import print
from rich.traceback import install

from core import LocalFileSystemScanner

load_dotenv()
install(show_locals=True)


def main() -> None:
    root = Path(__file__).resolve().parent
    scanner = LocalFileSystemScanner()
    tree = scanner.scan(root, sort_entries=True, allow_empty="allow")

    print(tree)
    print(f"Total Files: {tree.total_files}")
    print(f"Total Bytes: {tree.total_size}")
    print(f"Total levels: {tree.total_levels}")
    print(f"Total subdirectories: {tree.total_subdirectories}")


if __name__ == "__main__":
    main()
