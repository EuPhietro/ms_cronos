"""Modelos e colecoes semanticas do Core.

Este modulo concentra:
- contratos simples de entrada e saida do Core;
- referencias enxutas para site, drive, item e arquivo local;
- colecoes genericas reutilizaveis, separando leitura, mutabilidade e
  imutabilidade;
- snapshots do sistema de arquivos local usados no planejamento de uploads.

A ideia e que o restante do projeto dependa destes tipos em vez de consumir
objetos crus do SDK do Microsoft Graph.

O fluxo de diretorios possui duas representacoes:
- `FilesystemTree` descreve fielmente o que o scanner encontrou no disco;
- `StagingFilesystemTree` descreve o que foi preparado para envio.

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

# TODO(beta): evolucoes planejadas antes da primeira versao estavel.
# - definir a politica para links, sockets e outros itens locais nao regulares;
# - modelar uma arvore remota navegavel alem do resultado parcial atual;
# - decidir se a arvore de staging deve delegar os totais de `source`;
# - decidir se `RootUploadMode` integra o staging ou deve ser removido.
from __future__ import annotations

from abc import ABC, abstractmethod
from collections.abc import Awaitable, Callable, Iterable, Iterator, Sequence
from dataclasses import dataclass, field
from pathlib import Path, PurePosixPath
from typing import (
    ClassVar,
    Generic,
    Literal,
    Self,
    TypeVar,
    cast,
    overload,
)

from core.errors import (
    InvalidRemoteNameError,
    LocalFileNotReadableError,
    LocalPathError,
    LocalPathIsDirectoryError,
    LocalPathNotFoundError,
)
from core.urls import validate_remote_name, validate_remote_path

T = TypeVar("T")

ConflictBehavior = Literal["fail", "rename", "replace"]

RootUploadMode = Literal["include_root", "contents_only"]


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


class CollectionItem(ABC):
    """Marcador semantico para itens que circulam nas colecoes do Core.

    A classe nao adiciona comportamento por enquanto. Ela existe para deixar
    claro que `SharePointSite`, `DocumentLibrary`, `SharePointItem`, `LocalFile` e
    `FileUploadResult` sao modelos internos do Core e podem ser agrupados por
    colecoes semanticas.
    """


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


@dataclass
class TreeUploadResult(CollectionItem):
    """Registra o progresso materializado de um upload de arvore.

    A instancia e preenchida nivel por nivel e pode ser reutilizada como
    checkpoint em uma nova chamada de ``upload_tree`` dentro do mesmo processo.
    """

    source_root: Path | None = None
    library_id: str | None = None
    parent_item_id: str | None = None
    target_root: PurePosixPath | None = None
    staging_fingerprint: str | None = None
    remote_directories: dict[PurePosixPath, SharePointItem] = field(
        default_factory=dict
    )
    uploaded_files: dict[PurePosixPath, list[FileUploadResult]] = field(
        default_factory=dict
    )
    completed_levels: set[PurePosixPath] = field(default_factory=set)

    @property
    def total_uploaded_files(self) -> int:
        """Retorna a quantidade de arquivos confirmados pelo Graph."""
        return sum(len(files) for files in self.uploaded_files.values())


TreeUploadPhase = Literal[
    "preparing_directories",
    "uploading_files",
    "completed",
]


@dataclass(frozen=True)
class TreeUploadProgress:
    """Snapshot imutavel do progresso observado durante um upload de arvore."""

    phase: TreeUploadPhase
    completed_files: int
    total_files: int
    completed_levels: int
    total_levels: int
    current_path: PurePosixPath | None = None


TreeUploadProgressCallback = Callable[
    [TreeUploadProgress],
    Awaitable[None] | None,
]


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
        """Declara o armazenamento que sera definido pela subclasse concreta."""
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
        """Valida a existencia da raiz e a correspondencia de seu nome."""
        if not self.path.exists():
            raise LocalPathNotFoundError(
                f"O diretorio raiz local nao existe: '{self.path}'."
            )
        if not self.path.is_dir():
            raise LocalPathError(
                f"A raiz local deve representar um diretorio: '{self.path}'."
            )
        if self.name != self.path.name:
            raise LocalPathError(
                f"O nome da raiz deve ser '{self.path.name}', mas foi recebido "
                f"'{self.name}'."
            )

    @classmethod
    def from_uri(cls, path: str | Path) -> Self:
        """Constroi a raiz a partir de um caminho local."""
        resolved_path = Path(path)
        return cls(
            path=resolved_path,
            name=resolved_path.name,
        )


# Filesystem Models


@dataclass(frozen=True)
class LocalFolder(CollectionItem):
    """Referencia enxuta e opaca de um subdiretorio encontrado no disco.

    O modelo descreve somente o recurso local. Sua correspondencia com uma
    pasta remota pertence a `StagingFolder` e ao executor de upload.

    Exemplo:
        folder = LocalFolder.from_uri('/tmp/documentos/relatorios')
    """

    path: Path
    name: str

    def __post_init__(self):
        """Valida a existencia da pasta e a correspondencia de seu nome."""
        if not self.path.exists():
            raise LocalPathNotFoundError(
                f"O diretorio local nao existe: '{self.path}'."
            )
        if not self.path.is_dir():
            raise LocalPathError(
                f"O caminho local deve representar um diretorio: '{self.path}'."
            )
        if self.name != self.path.name:
            raise LocalPathError(
                f"O nome da pasta deve ser '{self.path.name}', mas foi recebido "
                f"'{self.name}'."
            )

    @classmethod
    def from_uri(cls, path: str | Path) -> Self:
        """Constroi uma referencia de pasta a partir de um caminho local."""
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
        """Valida metadados e aplica a politica para arquivos vazios."""
        if not self.path.exists():
            raise LocalPathNotFoundError(f"O arquivo local nao existe: '{self.path}'.")
        if not self.path.is_file():
            if self.path.is_dir():
                raise LocalPathIsDirectoryError(
                    f"O caminho local aponta para um diretorio: '{self.path}'."
                )
            raise LocalPathError(
                f"O caminho local nao representa um arquivo regular: '{self.path}'."
            )
        if self.name != self.path.name:
            raise LocalPathError(
                f"O nome do arquivo deve ser '{self.path.name}', mas foi recebido "
                f"'{self.name}'."
            )
        if self.extension != self.path.suffix:
            raise LocalPathError(
                f"A extensao do arquivo deve ser '{self.path.suffix}', mas foi "
                f"recebido '{self.extension}'."
            )
        if self.allow_empty not in {"allow", "deny"}:
            raise ValueError(
                "allow_empty deve ser 'allow' ou 'deny'; recebido: "
                f"{self.allow_empty!r}."
            )
        if self.size < 0:
            raise LocalFileNotReadableError(
                f"O tamanho do arquivo nao pode ser negativo: {self.size}."
            )
        if self.allow_empty == "deny" and self.size == 0:
            raise LocalFileNotReadableError(
                f"O arquivo '{self.path}' deve possuir ao menos um byte."
            )

    @classmethod
    def from_uri(cls, path: str | Path) -> Self:
        """Cria um `LocalFile` estrito a partir de um caminho local.

        A consulta a `stat()` exige que o recurso exista. As demais invariantes
        sao verificadas pelo `__post_init__`.
        """
        resolved_path = Path(path)
        if not resolved_path.exists():
            raise LocalPathNotFoundError(
                f"O arquivo local nao existe: '{resolved_path}'."
            )
        if not resolved_path.is_file():
            if resolved_path.is_dir():
                raise LocalPathIsDirectoryError(
                    f"O caminho local aponta para um diretorio: '{resolved_path}'."
                )
            raise LocalPathError(
                f"O caminho local nao representa um arquivo: '{resolved_path}'."
            )
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
        if not name.strip():
            raise LocalPathError("O novo nome local nao pode ser vazio.")

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

    def __post_init__(self) -> None:
        """Valida o diretorio e a pertinencia de seus filhos imediatos."""
        if not self.path.exists():
            raise LocalPathNotFoundError(
                f"O diretorio do nivel local nao existe: '{self.path}'."
            )
        if not self.path.is_dir():
            raise LocalPathError(
                f"O nivel local deve representar um diretorio: '{self.path}'."
            )
        for child in (*self.files, *self.folders):
            if child.path.parent != self.path:
                raise LocalPathError(
                    f"O recurso '{child.path}' nao e filho imediato do nivel "
                    f"'{self.path}'."
                )

    @property
    def total_size(self) -> int:
        """Soma os bytes dos arquivos diretamente contidos no nivel."""
        total = 0
        for file in self.files:
            total += file.size

        return total

    @property
    def total_files(self) -> int:
        """Retorna a quantidade de arquivos diretamente contidos."""
        return len(self.files)

    @property
    def total_directories(self) -> int:
        """Retorna a quantidade de subdiretorios diretamente contidos."""
        return len(self.folders)


@dataclass(frozen=True)
class DirectoryLevelCollection(FrozenCollection[DirectoryLevel]):
    """Colecao imutavel dos niveis locais na ordem produzida pelo scanner."""

    @property
    def total_size(self) -> int:
        """Soma os bytes registrados em todos os niveis da colecao."""
        return sum(level.total_size for level in self)

    @property
    def total_files(self) -> int:
        """Soma os arquivos registrados em todos os niveis."""
        return sum(level.total_files for level in self)


@dataclass(frozen=True, kw_only=True)
class FilesystemTree:
    """Snapshot plano de uma arvore de diretorios local.

    `root` define a origem da varredura. `levels` guarda cada diretorio
    encontrado, preferencialmente em ordem top-down, sem carregar em memoria o
    conteudo binario dos arquivos.

    Esta classe descreve o disco e nao deve conhecer o SharePoint.
    `StagingTreeBuilder` produz o `StagingFilesystemTree` consumido pelo upload
    de diretorios.
    """

    root: RootFolder
    levels: DirectoryLevelCollection

    def __post_init__(self):
        """Confirma que os recursos do snapshot ainda existem no disco."""
        root_path = self.root.path.resolve()
        if not self.levels:
            raise LocalPathError(
                f"A arvore local deve conter o nivel raiz '{root_path}'."
            )
        if self.levels[0].path.resolve() != root_path:
            raise LocalPathError(
                "O primeiro nivel da arvore deve representar a raiz local: "
                f"esperado '{root_path}', recebido '{self.levels[0].path}'."
            )
        for level in self.levels:
            try:
                level.path.resolve().relative_to(root_path)
            except ValueError as error:
                raise LocalPathError(
                    f"O nivel '{level.path}' esta fora da raiz '{root_path}'."
                ) from error

    @property
    def total_size(self) -> int:
        """Retorna o tamanho total dos arquivos registrados, em bytes."""
        return self.levels.total_size

    @property
    def total_files(self) -> int:
        """Retorna a quantidade total de arquivos registrados."""
        return self.levels.total_files

    @property
    def total_levels(self) -> int:
        """Retorna a quantidade de diretorios materializados como niveis."""
        return len(self.levels)

    @property
    def total_subdirectories(self) -> int:
        """Soma as referencias de subdiretorios imediatos de todos os niveis."""
        return sum(level.total_directories for level in self.levels)


# Os modelos locais acima descrevem o snapshot observado no disco. Os modelos
# abaixo acrescentam caminhos relativos e destinos para formar o plano
# semantico consumido pelo orquestrador de upload.


@dataclass(frozen=True)
class StagingFile(CollectionItem):
    """Representa um arquivo local preparado para envio.

    `source` preserva a referencia local original. `relative_path` posiciona o
    arquivo dentro da arvore escaneada. O `target_path` materializado deve ser
    derivado do `target_root` da arvore somado a esse caminho relativo;
    portanto, nao pode representar uma reorganizacao independente.
    `remote_name` registra o nome usado no SharePoint, e
    `conflict_behavior` define a estrategia para colisao.

    O model armazena apenas metadados. O conteudo binario continua no arquivo
    local e so deve ser aberto durante a execucao do upload. Seus fragmentos
    permanecem decodificados; a camada de URL aplicara percent-encoding.
    """

    source: LocalFile  # Arquivo local que sera enviado.
    relative_path: PurePosixPath  # Caminho relativo a raiz do `FilesystemTree`.
    remote_name: str  # Nome final planejado para o recurso remoto.
    conflict_behavior: ConflictBehavior = "fail"

    def __post_init__(self):
        """Valida a origem e os metadados essenciais do staging."""
        if not self.source.path.exists():
            raise LocalPathNotFoundError(
                "O arquivo local preparado para staging nao foi encontrado: "
                f"{self.source.path}."
            )
        if not self.source.path.is_file():
            if self.source.path.is_dir():
                raise LocalPathIsDirectoryError(
                    "O staging esperava um arquivo, mas o caminho local aponta "
                    f"para um diretorio: {self.source.path}."
                )
            raise LocalPathError(
                "O caminho local preparado para staging nao representa um "
                f"arquivo regular: {self.source.path}."
            )
        if self.relative_path.is_absolute():
            raise LocalPathError(
                "O relative_path do arquivo deve ser relativo a raiz da "
                f"arvore, mas foi recebido um caminho absoluto: "
                f"{self.relative_path}."
            )
        if ".." in self.relative_path.parts:
            raise LocalPathError(
                "O relative_path do arquivo nao pode sair da raiz da arvore "
                f"usando '..': {self.relative_path}."
            )
        if not self.remote_name.strip():
            raise InvalidRemoteNameError(
                "O nome remoto do arquivo nao pode ser vazio ou conter apenas "
                f"espacos. Origem local: {self.source.path}."
            )
        validate_remote_name(self.remote_name)


@dataclass(frozen=True)
class StagingFolder(CollectionItem):
    """Representa uma pasta local preparada para criacao remota.

    A origem permanece em `source`; `relative_path` preserva sua posicao na
    arvore local. O `target_path`, quando materializado, deve corresponder ao
    `target_root` da arvore somado a esse fragmento. `remote_name` preserva o
    nome planejado sem codificacao de URL. O model nao representa uma pasta
    remota ja criada.
    """

    source: LocalFolder
    relative_path: PurePosixPath
    remote_name: str

    def __post_init__(self):
        """Valida a origem e os metadados essenciais do staging da pasta."""
        if not self.source.path.exists():
            raise LocalPathNotFoundError(
                "A pasta local preparada para staging nao foi encontrada: "
                f"{self.source.path}."
            )
        if not self.source.path.is_dir():
            raise LocalPathError(
                "O staging esperava uma pasta, mas o caminho local nao "
                f"representa um diretorio: {self.source.path}."
            )
        if self.relative_path.is_absolute():
            raise LocalPathError(
                "O relative_path da pasta deve ser relativo a raiz da arvore, "
                f"mas foi recebido um caminho absoluto: {self.relative_path}."
            )
        if ".." in self.relative_path.parts:
            raise LocalPathError(
                "O relative_path da pasta nao pode sair da raiz da arvore "
                f"usando '..': {self.relative_path}."
            )
        if not self.remote_name.strip():
            raise InvalidRemoteNameError(
                "O nome remoto da pasta nao pode ser vazio ou conter apenas "
                f"espacos. Origem local: {self.source.path}."
            )
        validate_remote_name(self.remote_name)


@dataclass(frozen=True)
class StagingFileCollection(FrozenCollection[StagingFile]):
    """Colecao imutavel de arquivos preparados para upload."""


@dataclass(frozen=True)
class StagingFolderCollection(FrozenCollection[StagingFolder]):
    """Colecao imutavel de pastas preparadas para criacao remota."""


@dataclass(frozen=True)
class StagingDirectoryLevel(CollectionItem):
    """Representa o plano de envio do conteudo imediato de um diretorio.

    `source` aponta para o nivel observado pelo scanner. As colecoes de staging
    guardam somente os arquivos e subdiretorios imediatos desse nivel;
    descendentes mais profundos aparecem em outros itens da colecao de niveis.
    A estrutura remota preserva essa mesma hierarquia: `target_path` deve ser
    derivado do `target_root` da arvore e de `relative_path`.
    """

    source: DirectoryLevel
    relative_path: PurePosixPath
    staging_files: StagingFileCollection
    staging_folders: StagingFolderCollection

    def __post_init__(self):
        """Valida a natureza relativa dos fragmentos local e remoto."""
        if not self.source.path.exists():
            raise LocalPathNotFoundError(
                "O diretorio local representado pelo nivel de staging nao foi "
                f"encontrado: {self.source.path}."
            )
        if not self.source.path.is_dir():
            raise LocalPathError(
                "O nivel de staging deve representar um diretorio local, mas "
                f"recebeu: {self.source.path}."
            )
        if self.relative_path.is_absolute():
            raise LocalPathError(
                "O relative_path do nivel deve ser relativo a raiz da arvore, "
                f"mas foi recebido um caminho absoluto: {self.relative_path}."
            )
        if ".." in self.relative_path.parts:
            raise LocalPathError(
                "O relative_path do nivel nao pode sair da raiz da arvore "
                f"usando '..': {self.relative_path}."
            )


@dataclass(frozen=True)
class StagingDirectoryLevelCollection(FrozenCollection[StagingDirectoryLevel]):
    """Colecao imutavel dos niveis preparados para percurso top-down."""


@dataclass(frozen=True)
class StagingFilesystemTree:
    """Representa o plano completo de upload de uma arvore local.

    `source` preserva o snapshot local. `target_root` e um fragmento remoto
    relativo ao pai que o executor recebera posteriormente, nao um item remoto
    nem uma URL. O valor `"."` significa copiar os filhos da raiz local
    diretamente para esse pai; outro fragmento cria um prefixo para toda a
    estrutura.

    `levels` mantem o trabalho preparado em ordem top-down. Para qualquer
    nivel ou recurso, o destino e derivado por:

        target_path = target_root / relative_path

    A politica de conflito funciona como padrao para os arquivos da arvore.
    Nenhum fragmento deste model deve estar percent-encoded.
    """

    source: FilesystemTree
    levels: StagingDirectoryLevelCollection
    conflict_behavior: ConflictBehavior = "fail"
    target_root: PurePosixPath = PurePosixPath(".")

    def __post_init__(self):
        """Valida o fragmento remoto usado como raiz da arvore preparada."""
        validate_remote_path(self.target_root)
