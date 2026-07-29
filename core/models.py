"""Modelos e colecoes semanticas do Core.

Este modulo concentra:
- contratos simples de entrada e saida do Core;
- referencias enxutas para site, drive, item e arquivo local;
- colecoes genericas reutilizaveis, separando leitura, mutabilidade e
  imutabilidade;
- snapshots do sistema de arquivos local usados no planejamento de uploads.

A ideia e que o restante do projeto dependa destes tipos em vez de consumir
objetos crus do SDK do Microsoft Graph.

O fluxo planejado para diretorios possui duas representacoes:
- `FilesystemTree` descreve fielmente o que o scanner encontrou no disco;
- `StagingFilesystemTree` descrevera o que esta preparado para ser enviado.

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

# TODO(directory-upload): concluir o pipeline de preparacao de diretorios.
#
# Fluxo esperado:
#
#    LocalFilesystemScanner
#        -> FilesystemTree
#        -> conversor/builder de staging
#        -> StagingFilesystemTree
#        -> servico de upload
#
# O scanner deve somente inspecionar o disco e montar `FilesystemTree`. Ele nao
# deve conhecer bibliotecas, itens remotos, conflito de nomes ou chamadas ao
# Microsoft Graph.
#
# 1. Concluir o contrato do snapshot local:
#
#    - adicionar `relative_path: Path` a `DirectoryLevel`, ou documentar uma
#      operacao equivalente em `FilesystemTree`;
#    - calcular o tamanho de cada nivel somente com seus arquivos diretos;
#    - derivar tamanho e contadores totais a partir de `levels`;
#    - preservar diretorios vazios;
#    - manter arquivos vazios invalidos por padrao e representa-los somente
#      quando a politica explicita `allow_empty="allow"` for usada;
#    - preservar a ordem top-down produzida por `Path.walk()`.
#
# 2. Implementar os modelos de staging. Eles representam recursos locais ja
#    associados ao destino remoto e prontos para consumo pelo uploader:
#
#    @dataclass(frozen=True)
#    class StagingFile(CollectionItem):
#        source: LocalFile
#        target_path: str
#        conflict_behavior: ConflictBehavior
#
#    @dataclass(frozen=True)
#    class StagingFolder(CollectionItem):
#        source: LocalFolder
#        target_path: str
#
#    `target_path` deve ser um fragmento remoto relativo, nunca uma URL completa
#    do Graph nem um caminho absoluto do computador local.
#
#    @dataclass(frozen=True)
#    class StagingFileCollection(FrozenCollection[StagingFile]):
#        pass
#
#    @dataclass(frozen=True)
#    class StagingFolderCollection(FrozenCollection[StagingFolder]):
#        pass
#
#    @dataclass(frozen=True)
#    class StagingDirectoryLevel(CollectionItem):
#        source: DirectoryLevel
#        target_path: str
#        files: StagingFileCollection
#        folders: StagingFolderCollection
#
#    @dataclass(frozen=True)
#    class StagingDirectoryLevelCollection(
#        FrozenCollection[StagingDirectoryLevel]
#    ):
#        pass
#
#    @dataclass(frozen=True)
#    class StagingFilesystemTree:
#        source: FilesystemTree
#        library: DocumentLibrary
#        destination: SharePointItem
#        levels: StagingDirectoryLevelCollection
#
# 3. Implementar o conversor/builder que recebe `FilesystemTree`, biblioteca,
#    pasta remota inicial e politica de conflito, validando todos os caminhos
#    antes de produzir `StagingFilesystemTree`.
#
# 4. Decidir se `StagingFile` substituira `PreparedUpload` ou se o reutilizara
#    por composicao. Os dois modelos nao devem manter contratos concorrentes
#    para o mesmo arquivo preparado.
#
# 5. Manter `LocalFile`, `LocalFolder`, `DirectoryLevel` e `FilesystemTree`
#    independentes do SharePoint. Somente os modelos `Staging*` devem conhecer
#    destino remoto e politica de upload.

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

# BaseClasses


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


# Representa um item genérico de coleção
class CollectionItem(ABC):
    """Marcador semantico para itens que circulam nas colecoes do Core.

    A classe nao adiciona comportamento por enquanto. Ela existe para deixar
    claro que `SharePointSite`, `DocumentLibrary`, `SharePointItem`, `LocalFile` e
    `FileUploadResult` sao modelos internos do Core e podem ser agrupados por
    colecoes semanticas.
    """


# Remote DocumentLibraries  Models
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


@dataclass(frozen=True)
class RootFolder:
    """Representa o diretorio raiz escolhido para o snapshot local.

    `path` identifica a origem da varredura e `name` preserva o nome real do
    diretorio. O scanner usa essa raiz como base para calcular os caminhos
    relativos dos niveis descendentes.

    Exemplo:
        root = RootFolder.from_uri('/tmp/documentos')
    """

    path: Path
    name: str

    def __post_init__(self):
        assert self.path.exists(), (
            f"O caminho passado não é válido self.path.exists(): {self.path.exists()}"
        )
        assert self.name == self.path.name, (
            f"self.name deve corresponder ao nome real do Folder"
        )

    @classmethod
    def from_uri(cls, path: str | Path) -> Self:
        resolved_path = Path(path)
        return cls(
            path=resolved_path,
            name=resolved_path.name,
        )


# Filesystem Models


@dataclass(frozen=True)
class LocalFolder(CollectionItem):
    """Referencia enxuta de um subdiretorio encontrado no disco.

    O modelo descreve somente o recurso local. Sua correspondencia com uma
    pasta remota sera responsabilidade futura de `StagingFolder`.

    Exemplo:
        folder = LocalFolder.from_uri('/tmp/documentos/relatorios')
    """

    path: Path
    name: str

    def __post_init__(self):
        assert self.path.exists(), (
            f"O caminho passado não é válido self.path.exists(): {self.path.exists()}"
        )
        assert self.name == self.path.name, (
            f"self.name deve corresponder ao nome real do Folder"
        )

    @classmethod
    def from_uri(cls, path: str | Path) -> Self:
        resolved_path = Path(path)
        return cls(
            path=resolved_path,
            name=resolved_path.name,
        )


@dataclass(frozen=True)
class LocalFile(CollectionItem):
    """Modelo semantico para um arquivo local elegivel para upload.

    O caminho deve existir e o nome e a extensao devem corresponder ao recurso
    no disco. Arquivos vazios sao rejeitados por padrao e aceitos apenas quando
    `allow_empty="allow"` for informado explicitamente.

    Exemplo:
        local_file = LocalFile.from_uri('/tmp/curriculo.pdf')
    """

    path: Path
    name: str
    size: int
    extension: str
    allow_empty: Literal["allow", "deny"] = "deny"

    def __post_init__(self):
        assert self.path.exists(), (
            f"O caminho passado não é válido self.path.exists(): {self.path.exists()}"
        )
        assert self.name == self.path.name, (
            "self.name deve corresponder ao nome real do arquivo"
        )
        assert self.extension == self.path.suffix, (
            "self.extension deve corresponder a extensão real"
        )
        if self.allow_empty == "deny":
            assert self.size > 0, "self.size deve ter conteúdo"
        if self.allow_empty == "allow":
            assert self.size >= 0

    @classmethod
    def from_uri(cls, path: str | Path) -> Self:
        """Cria um `LocalFile` estrito a partir de um caminho local.

        A consulta a `stat()` exige que o recurso exista. As demais invariantes
        sao verificadas pelo `__post_init__`.
        """
        resolved_path = Path(path)
        suffix = resolved_path.suffix
        return cls(
            path=resolved_path,
            name=resolved_path.name,
            extension=suffix,
            size=resolved_path.stat().st_size,
        )

    def rename_on_disk(self, name: str) -> Self:
        """Renomeia o arquivo no disco e devolve uma referencia atualizada.

        Esta operacao altera o filesystem local. Ela nao serve apenas para
        escolher um nome remoto diferente.
        """
        assert len(name) > 0

        parent = self.path.parent
        new_path = parent / name
        new_path = self.path.rename(new_path)
        return type(self)(new_path, name, self.size, new_path.suffix, self.allow_empty)


@dataclass(frozen=True)
class LocalFolderCollection(FrozenCollection[LocalFolder]):
    """Colecao imutavel de subdiretorios pertencentes a um nivel local.

    Exemplo:
        folders = LocalFolderCollection.from_collection([folder])
    """


@dataclass(frozen=True)
class LocalFileCollection(FrozenCollection[LocalFile]):
    """Colecao imutavel de arquivos pertencentes a um nivel local.

    Exemplo:
        files = LocalFileCollection.from_collection([local_file])
        assert files.first() == local_file
    """


@dataclass(frozen=True, kw_only=True)
class DirectoryLevel(CollectionItem):
    """Representa um diretorio e seu conteudo local imediato.

    O nivel corresponde a uma tupla produzida por `Path.walk()`: `path`
    identifica o diretorio atual, `files` contem somente seus arquivos diretos
    e `folders` contem somente seus subdiretorios diretos. Descendentes mais
    profundos aparecem como outros itens de `DirectoryLevelCollection`.
    """

    path: Path  # Caminho absoluto do diretorio representado por este nivel.
    files: LocalFileCollection  # Arquivos diretamente contidos no nivel.
    folders: LocalFolderCollection  # Subdiretorios diretamente contidos.

    @property
    def total_size(self) -> int:
        total = 0
        for file in self.files:
            total += file.size

        return total

    @property
    def total_files(self) -> int:
        return len(self.files)

    @property
    def total_directories(self) -> int:
        return len(self.folders)


@dataclass(frozen=True)
class DirectoryLevelCollection(FrozenCollection[DirectoryLevel]):
    """Colecao imutavel dos niveis locais na ordem produzida pelo scanner."""

    @property
    def total_size(self) -> int:
        return sum(level.total_size for level in self)

    @property
    def total_files(self) -> int:
        return sum(level.total_files for level in self)


@dataclass(frozen=True, kw_only=True)
class FilesystemTree:
    """Snapshot plano de uma arvore de diretorios local.

    `root` define a origem da varredura. `levels` guarda cada diretorio
    encontrado, preferencialmente em ordem top-down, sem carregar em memoria o
    conteudo binario dos arquivos.

    Esta classe descreve o disco e nao deve conhecer o SharePoint. Um conversor
    futuro produzira `StagingFilesystemTree`, que sera o contrato de entrada do
    upload de diretorios.
    """

    root: RootFolder
    levels: DirectoryLevelCollection

    def __post_init__(self):

        assert self.root.path.exists()

        for level in self.levels:
            assert level.path.exists()
            for dir in level.folders:
                assert dir.path.exists()
            for file in level.files:
                assert file.path.exists()

    @property
    def total_size(self) -> int:
        return self.levels.total_size

    @property
    def total_files(self) -> int:
        return self.levels.total_files

    @property
    def total_levels(self) -> int:
        return len(self.levels)

    @property
    def total_subdirectories(self) -> int:
        return sum(level.total_directories for level in self.levels)


# Os modelos locais acima registram o snapshot do disco. Os modelos `Staging*`
# descritos no TODO do modulo formarao, posteriormente, o plano semantico de
# upload para o SharePoint.


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


@dataclass
class PreparedUploadCollection(MutableCollection[PreparedUpload]):
    """Colecao mutavel de uploads pequenos ja preparados para envio.

    Exemplo:
        staged = PreparedUploadCollection(_slots=[staging_content])
        staged.clear()
        assert staged.is_empty
    """
