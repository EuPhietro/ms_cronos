"""Helpers pequenos e puros reutilizados pelo Core.

Este modulo concentra utilitarios que nao dependem do SDK do Graph e nao
conhecem regras de negocio maiores. Eles existem para evitar repeticao em
fluxos de naming e preparacao de dados.
"""

from pathlib import PurePath
from uuid import uuid4


def rename_with_uuid(name: str) -> str:
    """Gera um nome alternativo anexando um UUID sem perder a extensao."""
    id = uuid4()
    if not name.strip():
        return str(id)

    # O nome e tratado como path puro apenas para separar stem e extensoes sem
    # tocar no sistema de arquivos local.
    path = PurePath(name)
    suffix = "".join(path.suffixes)
    stem = path.name[: -len(suffix)] if suffix else path.name

    if not stem:
        return f"{id}{suffix}"
    return f"{stem}-{id}{suffix}"
