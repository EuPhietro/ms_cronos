"""Modelos e colecoes semanticas do Core.

Este modulo concentra:
- contratos simples de entrada e saida do Core;
- referencias enxutas para site, drive, item e arquivo local;
- colecoes genericas reutilizaveis, separando leitura, mutabilidade e
  imutabilidade.

A ideia e que o restante do projeto dependa destes tipos em vez de consumir
objetos crus do SDK do Microsoft Graph.

Exemplo basico:
    credentials = GraphCredentials(
        client_id="app-id",
        client_secret="secret",
        tenant_id="tenant-id",
    )

    site = SharePointSite(
        id="tenant.sharepoint.com,site-guid,web-guid",
        name="RHConecta",
        display_name="RH Conecta",
        web_url="https://tenant.sharepoint.com/sites/RHConecta",
    )

    sites = SharePointSiteCollection.from_collection([site])
    assert not sites.is_empty
    assert sites.first() == site
"""

from __future__ import annotations

from abc import ABC, abstractmethod
from collections.abc import Callable, Iterable, Iterator, Sequence
from dataclasses import dataclass, field
from pathlib import Path
from typing import (
    ClassVar,
    Generic,
    Literal,
    Self,
    TypeVar,
    cast,
    overload,
)

# `T` representa o tipo de item armazenado por uma colecao generica.
T = TypeVar("T")

# Politica de conflito aceita pelas operacoes de criacao e upload.
ConflictBehavior = Literal["fail", "rename", "replace"]


# `CollectionItem` permanece como classe base semantica para identificar os
# modelos que podem viver dentro das colecoes do Core.
class CollectionItem(ABC):
    """Marcador semantico para itens que circulam nas colecoes do Core.

    A classe nao adiciona comportamento por enquanto. Ela existe para deixar
    claro que `SharePointSite`, `DocumentLibrary`, `SharePointItem`, `LocalFile` e
    `FileUploadResult` sao modelos internos do Core e podem ser agrupados por
    colecoes semanticas.
    """


@dataclass(frozen=True)
class GraphCredentials:
    """
    Contrato de entrada das credenciais usadas para autenticar no Graph.

    Exemplo:
        credentials = GraphCredentials(
            client_id='app-id',
            client_secret='secret',
            tenant_id='tenant-id',
        )
    """

    client_id: str
    client_secret: str
    tenant_id: str


@dataclass(frozen=True)
class SharePointSite(CollectionItem):
    """
    Referencia enxuta de um site do SharePoint resolvido pelo Core.

    Exemplo:
        site = SharePointSite(
            id='tenant.sharepoint.com,site-guid,web-guid',
            name='RHConecta',
            display_name='RH Conecta',
            web_url='https://tenant.sharepoint.com/sites/RHConecta',
        )
    """

    id: str
    name: str | None = None
    display_name: str | None = None
    web_url: str | None = None


@dataclass(frozen=True)
class DocumentLibrary(CollectionItem):
    """
    Referencia enxuta de um drive do SharePoint.

    Exemplo:
        drive = DocumentLibrary(
            id='b!abc123',
            name='Documents',
            web_url='https://tenant.sharepoint.com/sites/RHConecta/Shared%20Documents',
            drive_type='documentLibrary',
        )
    """

    id: str
    name: str | None = None
    web_url: str | None = None
    drive_type: str | None = None


@dataclass(frozen=True)
class SharePointItem(CollectionItem):
    """
    Referencia enxuta de um item de drive, seja arquivo ou pasta.

    Exemplo:
        item = SharePointItem(
            id='01ABCDEF',
            name='Curriculos',
            web_url='https://tenant.sharepoint.com/sites/RHConecta/Shared%20Documents/Curriculos',
            is_folder=True,
            is_file=False,
            size=0,
        )
    """

    id: str
    name: str | None = None
    web_url: str | None = None
    is_folder: bool = False
    is_file: bool = False
    size: int | None = None


@dataclass(frozen=True)
class LocalFile(CollectionItem):
    """
    Modelo semantico para um arquivo local que pode ser enviado ao SharePoint.

    Exemplo:
        local_file = LocalFile.from_uri('/tmp/curriculo.pdf')
    """

    path: Path
    name: str
    extension: str | None = None
    size: int | None = None

    @classmethod
    def from_uri(cls, path: str | Path) -> LocalFile:
        """Cria um `LocalFile` a partir de um caminho local `str` ou `Path`.

        Este construtor e permissivo: se o caminho ainda nao existir, `size`
        permanece `None`. Validacoes mais rigidas ficam no parser ou no fluxo
        de upload.
        """
        resolved_path = Path(path)
        suffix = resolved_path.suffix or None
        return cls(
            path=resolved_path,
            name=resolved_path.name,
            extension=suffix,
            size=resolved_path.stat().st_size if resolved_path.exists() else None,
        )

    def rename(self, name: str) -> Self:
        """Retorna uma nova referencia local com outro nome sem mover o
        arquivo."""
        return type(self)(self.path, name, self.extension, self.size)


@dataclass(frozen=True)
class PreparedUpload(CollectionItem):
    """Representa um upload pequeno ja preparado para envio.

    O objeto guarda:
    - o arquivo local de origem;
    - o fragmento de path Graph que identifica o recurso de criacao;
    - a estrategia semantica de conflito pedida pelo chamador.
    """

    file: LocalFile
    target_path: str
    conflict_behavior: ConflictBehavior = "fail"


@dataclass(frozen=True)
class FileUploadResult(CollectionItem):
    """Representa o resultado semantico de um upload concluido.

    Exemplo:
        result = FileUploadResult(
            item=drive_item,
            source_path=Path('/tmp/curriculo.pdf'),
            remote_name='curriculo.pdf',
            conflict_behavior='replace',
        )
    """

    item: SharePointItem
    source_path: Path
    remote_name: str
    conflict_behavior: ConflictBehavior


class Collection_(Generic[T], ABC):
    """Contrato base de leitura para colecoes do Core.

    Esta classe nao decide se a colecao e mutavel ou imutavel; ela apenas
    oferece o comportamento comum de leitura, navegacao e inspecao.

    Exemplo:
        collection = FrozenCollection(_slots=('a', 'b'))
        assert collection.counter == 2
        assert collection.first() == 'a'
    """

    # Cada subclasse concreta informa como uma sequencia deve ser materializada
    # internamente. O tipo fica propositalmente amplo porque `ClassVar` nao
    # deve depender do `T` generico da instancia.
    _storage_factory: ClassVar[Callable[[Iterable[object]], Sequence[object]]]

    def __init__(self) -> None:
        # `_slots` e o armazenamento interno compartilhado por todas as
        # colecoes. Cada subclasse define se ele sera tupla ou lista.
        self._slots: Sequence[T]

    @classmethod
    @abstractmethod
    def from_collection(cls, collection: Sequence[T]) -> Self:
        """Reconstrói a colecao concreta a partir de uma sequencia de itens."""
        ...

    @property
    def counter(self) -> int:
        """Retorna a quantidade de itens armazenados."""
        return len(self._slots)

    @property
    def is_empty(self) -> bool:
        """Indica se a colecao nao possui itens."""
        return len(self._slots) == 0

    def first(self) -> T | None:
        """Retorna o primeiro item ou `None` quando a colecao estiver vazia."""
        return self._slots[0] if not self.is_empty else None

    def to_list(self) -> list[T]:
        """Devolve uma copia rasa da colecao como lista comum do Python."""
        return list(self._slots)

    def __iter__(self) -> Iterator[T]:
        """Permite iterar diretamente sobre a colecao."""
        return self._slots.__iter__()

    def __len__(self) -> int:
        """Permite usar `len(colecao)`."""
        return len(self._slots)

    def __bool__(self) -> bool:
        """Colecoes vazias sao avaliadas como `False`."""
        return not self.is_empty

    @overload
    def __getitem__(self, key: int) -> T: ...

    @overload
    def __getitem__(self, key: slice) -> Sequence[T]: ...

    def __getitem__(self, key: int | slice) -> T | Sequence[T]:
        """Permite acessar itens por indice ou slice."""
        return self._slots[key]

    def __contains__(self, item: object) -> bool:
        """Permite testar pertinencia com o operador `in`."""
        return item in self._slots


@dataclass(frozen=True)
class FrozenCollection(Collection_[T]):
    """
    Colecao imutavel usada como contrato publico de retorno do Core.

    Toda operacao que alteraria o conteudo devolve uma nova instancia em vez de
    modificar a colecao atual.

    Exemplo:
        drives = FrozenCollection(_slots=('Documents',))
        updated = drives.add('Curriculos')
        assert drives.counter == 1
        assert updated.counter == 2
    """

    _storage_factory = tuple
    _slots: tuple[T, ...] = field(  # pyright: ignore[reportIncompatibleVariableOverride]
        default_factory=tuple,
        init=True,
    )

    @classmethod
    def from_collection(cls, collection: Sequence[T]) -> Self:
        """Materializa uma sequencia como colecao imutavel do mesmo tipo."""
        if not collection:
            return cls()
        _slots_ = cast(tuple[T, ...], cls._storage_factory(collection))
        return cls(_slots=_slots_)

    def add(self, item: T) -> Self:
        """Retorna uma nova colecao contendo o item informado ao final."""
        _new_slots: tuple[T, ...] = (*self._slots, item)
        return type(self).from_collection(_new_slots)

    def extend(self, *items: T) -> Self:
        """Retorna uma nova colecao contendo os itens atuais e os novos."""
        _new_slots: tuple[T, ...] = (*self._slots, *items)

        return type(self).from_collection(_new_slots)

    def remove(self, item: T) -> Self:
        """Retorna uma nova colecao sem a primeira ocorrencia do item."""
        updated_slots = list(self._slots)
        updated_slots.remove(item)
        _new_slots: tuple[T, ...] = tuple(updated_slots)

        return type(self).from_collection(_new_slots)

    def clear(self) -> Self:
        """Retorna uma nova colecao vazia do mesmo tipo."""
        return type(self).from_collection(())


@dataclass
class MutableCollection(Collection_[T]):
    """Colecao editavel usada em cenarios de montagem e transformacao.

    Exemplo:
        files = MutableCollection(_slots=['a.txt'])
        files.add('b.txt')
        assert files.counter == 2
    """

    _storage_factory = list
    _slots: list[T] = field(  # pyright: ignore[reportIncompatibleVariableOverride]
        default_factory=list,
    )

    @classmethod
    def from_collection(cls, collection: Sequence[T]) -> Self:
        """Materializa uma sequencia como colecao mutavel do mesmo tipo."""
        if not collection:
            return cls()
        _slots_ = cast(list[T], cls._storage_factory(collection))

        return cls(_slots=_slots_)

    def add(self, item: T) -> None:
        """Adiciona um item ao final da colecao atual."""
        self._slots.append(item)

    def extend(self, *items: T) -> None:
        """Adiciona varios itens ao final da colecao atual."""
        self._slots.extend(items)

    def remove(self, item: T) -> None:
        """Remove a primeira ocorrencia do item informado."""
        self._slots.remove(item)

    def clear(self) -> None:
        """Remove todos os itens da colecao atual."""
        self._slots.clear()


@dataclass(frozen=True)
class SharePointSiteCollection(FrozenCollection[SharePointSite]):
    """Colecao imutavel de referencias de sites resolvidos pelo Core.

    Exemplo:
        sites = SharePointSiteCollection(_slots=(site_ref,))
        first_site = sites.first()
    """


@dataclass(frozen=True)
class DocumentLibraryCollection(FrozenCollection[DocumentLibrary]):
    """Colecao imutavel de referencias de drives do SharePoint.

    Exemplo:
        drives = DocumentLibraryCollection(_slots=(drive_ref,))
        assert drives.counter == 1
    """


@dataclass(frozen=True)
class SharePointItemCollection(FrozenCollection[SharePointItem]):
    """Colecao imutavel de arquivos e pastas retornados de um drive.

    Exemplo:
        items = SharePointItemCollection(_slots=(drive_item_ref,))
        assert drive_item_ref in items
    """


@dataclass
class LocalFileCollection(MutableCollection[LocalFile]):
    """Colecao mutavel de arquivos locais preparados para processamento.

    Exemplo:
        files = LocalFileCollection(_slots=[local_file])
        files.clear()
        assert files.is_empty
    """


@dataclass
class PreparedUploadCollection(MutableCollection[PreparedUpload]):
    """Colecao mutavel de uploads pequenos ja preparados para envio.

    Exemplo:
        staged = PreparedUploadCollection(_slots=[staging_content])
        staged.clear()
        assert staged.is_empty
    """
