"""Transformacao pura de uma arvore local em uma arvore de staging.

O modulo recebe o snapshot produzido por ``LocalFileSystemScanner`` e adapta
cada recurso local para os contratos ``Staging*``. Ele nao cria diretorios no
SharePoint, nao abre arquivos, nao carrega bytes e nao monta URLs do Graph.

Fluxo atual:

    FilesystemTree
        -> StagingDirectoryLevelCollection
        -> StagingDirectoryLevel
        -> StagingFileCollection e StagingFolderCollection
        -> StagingFilesystemTree

Os caminhos absolutos permanecem somente nos models locais de origem. Para o
plano de upload, cada caminho e convertido em um ``PurePosixPath`` relativo a
raiz do ``FilesystemTree``. A futura execucao remota combinara esses caminhos
com o ``target_root`` da arvore de staging, que representa um fragmento sob o
item remoto escolhido como pai e nunca uma URL completa do Graph.
"""

from pathlib import Path, PurePosixPath

from core.models import (
    ConflictBehavior,
    DirectoryLevel,
    DirectoryLevelCollection,
    FilesystemTree,
    LocalFile,
    LocalFileCollection,
    LocalFolder,
    LocalFolderCollection,
    StagingDirectoryLevel,
    StagingDirectoryLevelCollection,
    StagingFile,
    StagingFileCollection,
    StagingFilesystemTree,
    StagingFolder,
    StagingFolderCollection,
)
from core.urls import validate_remote_path


class StagingTreeBuilder:
    """Constroi o plano semantico de upload a partir de um snapshot local.

    A classe apenas transforma models. A ordem dos niveis recebidos e
    preservada para que o futuro executor possa criar os diretorios de cima
    para baixo antes de enviar os arquivos.
    """

    def build_staging_tree(
        self,
        tree: FilesystemTree,
        conflict_behavior: ConflictBehavior = "fail",
        target_root: PurePosixPath = PurePosixPath("."),
    ) -> StagingFilesystemTree:
        """Converte uma ``FilesystemTree`` em ``StagingFilesystemTree``.

        O caminho da pasta raiz local e usado somente como ancora para calcular
        os caminhos relativos. A politica de conflito e propagada para cada
        ``StagingFile`` e registrada tambem na arvore resultante. ``target_root``
        define o prefixo remoto comum a todos os recursos; ``.`` representa o
        proprio item pai que sera recebido posteriormente pelo executor.

        Args:
            tree: Snapshot completo da arvore encontrada no disco.
            conflict_behavior: Politica aplicada aos arquivos preparados.
            target_root: Fragmento remoto relativo usado como prefixo comum da
                arvore preparada. O valor ``.`` nao cria um prefixo adicional.

        Returns:
            Arvore de staging com os niveis convertidos na ordem original.

        Raises:
            InvalidRemoteNameError: Se ``target_root`` contiver um segmento
                remoto invalido.
        """
        validate_remote_path(target_root)
        return StagingFilesystemTree(
            source=tree,
            levels=self._build_staging_directory_level_collection(
                tree.root.path, source=tree.levels, conflict_behavior=conflict_behavior
            ),
            target_root=target_root,
            conflict_behavior=conflict_behavior,
        )

    def _build_staging_directory_level_collection(
        self,
        root: Path,
        source: DirectoryLevelCollection,
        conflict_behavior: ConflictBehavior = "fail",
    ) -> StagingDirectoryLevelCollection:
        """Converte todos os niveis locais em uma colecao de staging.

        Args:
            root: Caminho absoluto usado como base da relativizacao.
            source: Niveis produzidos pelo scanner, em ordem top-down.
            conflict_behavior: Politica propagada aos arquivos de cada nivel.

        Returns:
            Colecao imutavel que preserva a ordem dos niveis de origem.
        """

        return StagingDirectoryLevelCollection.from_collection(
            [
                self._build_staging_directory_level(
                    source=level, root=root, conflict_behavior=conflict_behavior
                )
                for level in source
            ]
        )

    def _build_staging_directory_level(
        self,
        root: Path,
        source: DirectoryLevel,
        conflict_behavior: ConflictBehavior = "fail",
    ) -> StagingDirectoryLevel:
        """Prepara o conteudo imediato de um diretorio local.

        O nivel agrega somente os arquivos e subdiretorios diretamente contidos
        em ``source``. Descendentes mais profundos permanecem representados por
        outros niveis da colecao principal.

        Args:
            root: Raiz local usada para calcular os caminhos relativos.
            source: Nivel local que sera adaptado.
            conflict_behavior: Politica propagada aos arquivos imediatos.

        Returns:
            Nivel de staging com arquivos e pastas imediatos preparados.
        """

        return StagingDirectoryLevel(
            source=source,
            relative_path=self._build_relative_path(root, source.path),
            staging_files=self._build_staging_file_collection(
                root, source.files, conflict_behavior=conflict_behavior
            ),
            staging_folders=self._build_staging_folder_collection(
                root=root, source=source.folders
            ),
        )

    def _build_staging_folder_collection(
        self,
        root: Path,
        source: LocalFolderCollection,
    ) -> StagingFolderCollection:
        """Prepara as referencias das subpastas imediatas de um nivel.

        Cada pasta conserva sua origem local, recebe um caminho relativo a
        arvore e, nesta primeira versao, reutiliza o nome local como nome
        remoto planejado.
        """
        return StagingFolderCollection.from_collection(
            [
                self._build_staging_folder(
                    source=folder,
                    relative_path=self._build_relative_path(
                        root=root, segment=folder.path
                    ),
                    remote_name=folder.name,
                )
                for folder in source
            ]
        )

    def _build_staging_file_collection(
        self,
        root: Path,
        source: LocalFileCollection,
        conflict_behavior: ConflictBehavior = "fail",
    ) -> StagingFileCollection:
        """Prepara os arquivos imediatos de um nivel para upload posterior.

        O metodo nao le o conteudo dos arquivos. Ele registra apenas a origem,
        a posicao relativa, o nome remoto planejado e a politica de conflito.
        """
        return StagingFileCollection.from_collection(
            [
                self._build_staging_file(
                    source=file,
                    relative_path=self._build_relative_path(
                        root=root, segment=file.path
                    ),
                    remote_name=file.name,
                    conflict_behavior=conflict_behavior,
                )
                for file in source
            ]
        )

    def _build_staging_folder(
        self, source: LocalFolder, relative_path: PurePosixPath, remote_name: str
    ) -> StagingFolder:
        """Materializa o model de staging de uma unica pasta local."""

        return StagingFolder(
            source=source, relative_path=relative_path, remote_name=remote_name
        )

    def _build_staging_file(
        self,
        source: LocalFile,
        relative_path: PurePosixPath,
        remote_name: str,
        conflict_behavior: ConflictBehavior = "fail",
    ) -> StagingFile:
        """Materializa o model de staging de um unico arquivo local."""

        return StagingFile(
            source=source,
            relative_path=relative_path,
            remote_name=remote_name,
            conflict_behavior=conflict_behavior,
        )

    def _build_relative_path(self, root: Path, segment: Path) -> PurePosixPath:
        """Calcula a posicao logica de um recurso dentro da arvore.

        ``Path.relative_to`` garante que ``segment`` esteja contido em
        ``root`` e levanta ``ValueError`` quando essa relacao nao existe. O
        resultado e convertido para ``PurePosixPath`` porque sera usado como
        caminho logico e remoto, sem acesso ao filesystem.

        Args:
            root: Caminho absoluto da raiz escaneada.
            segment: Caminho absoluto do recurso local.

        Returns:
            Caminho POSIX relativo a raiz; a propria raiz resulta em ``.``.

        Raises:
            ValueError: Se o recurso nao estiver contido na raiz informada.
        """
        return PurePosixPath(segment.relative_to(root))
