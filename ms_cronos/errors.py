"""Hierarquia de erros semanticos do Core.

As camadas superiores devem capturar erros deste modulo, nao erros crus do SDK
do Microsoft Graph. Isso mantem o restante do projeto desacoplado de detalhes
de infraestrutura.

Exemplo:
    try:
        site = await sharepoint.resolve_site(url)
    except SiteResolutionError:
        ...
    except MSCronosError:
        ...
"""


class MSCronosError(Exception):
    """Erro base do projeto.

    Use como raiz para qualquer excecao publica do Core. O restante do sistema
    deve conseguir capturar MSCronosError sem conhecer detalhes internos do
    SDK.
    """

    def __init__(self, *args: object) -> None:
        super().__init__(*args)


# ---- CONFIGURACOES E AUTENTICACAO ----


class GraphConfigurationError(MSCronosError):
    """Use quando a configuracao do cliente Graph estiver invalida antes de
    qualquer chamada externa, como credenciais ausentes, vazias ou scopes
    inconsistentes."""

    def __init__(self, *args: object) -> None:
        super().__init__(*args)


class GraphAuthenticationError(MSCronosError):
    """Use quando a autenticacao no Microsoft Graph falhar, mesmo com a
    configuracao formalmente valida, como client secret incorreto ou tenant
    invalido."""

    def __init__(self, *args: object) -> None:
        super().__init__(*args)


# ---- ENTRADA E VALIDACAO ----


class SharePointUrlError(MSCronosError):
    """Use quando a URL do SharePoint fornecida nao puder ser validada ou
    convertida para a rota esperada pelo Microsoft Graph."""

    def __init__(self, *args: object) -> None:
        super().__init__(*args)


class InvalidConflictBehaviorError(MSCronosError):
    """Use quando o valor de conflict behavior nao estiver entre as opcoes
    suportadas pelo Core, como replace, rename ou fail."""

    def __init__(self, *args: object) -> None:
        super().__init__(*args)


class InvalidRemoteNameError(MSCronosError):
    """Use quando o nome remoto do arquivo ou da pasta for invalido antes da
    chamada ao Graph, como nome vazio ou contendo separadores de caminho."""

    def __init__(self, *args: object) -> None:
        super().__init__(*args)


class LocalPathError(MSCronosError):
    """Classe base para erros de caminho local.

    Use para agrupar falhas relacionadas ao arquivo de origem antes do upload.
    """

    def __init__(self, *args: object) -> None:
        super().__init__(*args)


class LocalPathNotFoundError(LocalPathError):
    """Use quando o caminho local informado nao existir no disco."""

    def __init__(self, *args: object) -> None:
        super().__init__(*args)


class LocalPathIsDirectoryError(LocalPathError):
    """Use quando o caminho local existir, mas apontar para um diretorio onde o
    Core esperava um arquivo."""

    def __init__(self, *args: object) -> None:
        super().__init__(*args)


class LocalFileNotReadableError(LocalPathError):
    """Use quando o arquivo existir, mas nao puder ser lido pelo processo
    atual, como em casos de permissao ou bloqueio de acesso."""

    def __init__(self, *args: object) -> None:
        super().__init__(*args)


# ---- SITES, DRIVES E ITENS ----


class SiteResolutionError(MSCronosError):
    """Use quando o Core nao conseguir resolver um site SharePoint a partir da
    URL informada, mesmo com a URL formalmente valida."""

    def __init__(self, *args: object) -> None:
        super().__init__(*args)


class DriveNotFoundError(MSCronosError):
    """Use quando um drive solicitado por identificador nao existir ou nao
    puder ser localizado no contexto atual."""

    def __init__(self, *args: object) -> None:
        super().__init__(*args)


class DefaultDriveNotFoundError(MSCronosError):
    """Use quando um site for resolvido corretamente, mas o drive padrao nao
    puder ser obtido ou nao estiver acessivel."""

    def __init__(self, *args: object) -> None:
        super().__init__(*args)


class DriveItemNotFoundError(MSCronosError):
    """Use quando um item de drive esperado nao for encontrado, seja arquivo ou
    pasta."""

    def __init__(self, *args: object) -> None:
        super().__init__(*args)


class FolderNotFoundError(DriveItemNotFoundError):
    """Use quando o item esperado especificamente como pasta nao for
    encontrado."""

    def __init__(self, *args: object) -> None:
        super().__init__(*args)


class NotAFolderError(MSCronosError):
    """Use quando o item existir, mas o fluxo exigir uma pasta e o recurso
    retornado for outro tipo, como arquivo."""

    def __init__(self, *args: object) -> None:
        super().__init__(*args)


class NotAChildError(MSCronosError):
    """Use quando um item retornado nao puder ser tratado como filho valido
    dentro do contexto atual de navegacao ou upload."""

    def __init__(self, *args: object) -> None:
        super().__init__(*args)


class NotAFileError(MSCronosError):
    """Use quando o item existir, mas o fluxo exigir um arquivo e o recurso
    retornado for outro tipo, como pasta."""

    def __init__(self, *args: object) -> None:
        super().__init__(*args)


class FileAlreadyExistError(MSCronosError):
    """Use quando o fluxo de upload pedir falha em caso de conflito e ja
    existir um arquivo remoto com o mesmo nome no destino."""

    def __init__(self, *args: object) -> None:
        super().__init__(*args)


# ---- ERROS DE REQUISICAO GRAPH ----


class FileVeryLargeError(MSCronosError):
    """Use quando um arquivo exceder o limite aceito pelo fluxo de upload
    pequeno e precisar ser redirecionado para upload com sessao."""

    def __init__(self, *args: object) -> None:
        super().__init__(*args)


class GraphRequestError(MSCronosError):
    """Erro generico para falhas de requisicao ao Microsoft Graph.

    Use quando houve resposta de erro ou falha operacional remota que nao se
    encaixa em um subtipo mais especifico.
    """

    def __init__(self, *args: object) -> None:
        super().__init__(*args)


class GraphPermissionError(GraphRequestError):
    """Use quando a operacao falhar por falta de permissao, autorizacao
    insuficiente ou escopo inadequado no Microsoft Graph."""

    def __init__(self, *args: object) -> None:
        super().__init__(*args)


class GraphResourceConflictError(GraphRequestError):
    """Use quando o Graph retornar conflito de recurso, como nome ja existente
    ou estado remoto incompatível com a operacao solicitada."""

    def __init__(self, *args: object) -> None:
        super().__init__(*args)


class GraphResponseError(GraphRequestError):
    """Use quando a resposta do Graph vier incompleta, inconsistente ou sem os
    dados minimos esperados para o contrato do Core."""

    def __init__(self, *args: object) -> None:
        super().__init__(*args)


class GraphTransportError(GraphRequestError):
    """Use quando a conexao com o Microsoft Graph falhar antes de uma resposta."""

    def __init__(self, *args: object) -> None:
        super().__init__(*args)


class CheckpointError(MSCronosError):
    """Erro base para leitura, gravacao ou validacao de checkpoints."""


class CheckpointFormatError(CheckpointError):
    """Use quando um checkpoint nao possuir o schema esperado pelo Core."""


class CheckpointMismatchError(CheckpointError):
    """Use quando o checkpoint pertencer a outra origem ou destino remoto."""


# ---- ERROS DE CRIAÇÃO DE RECURSOS ----


class FailedWhenCreateDriveItemError(MSCronosError):
    """Use quando a operacao de criacao nao devolver um `DriveItem` valido para
    o Core seguir com parse e navegacao."""

    def __init__(self, *args: object) -> None:
        super().__init__(*args)


# ---- ERROS DE UPLOAD ----


class UploadError(MSCronosError):
    """Classe base para erros de upload.

    Use para agrupar falhas ocorridas durante o envio de arquivos ao
    SharePoint.
    """

    def __init__(self, *args: object) -> None:
        super().__init__(*args)


class TreeUploadError(UploadError):
    """Use como erro base da materializacao e do upload de uma arvore."""

    def __init__(self, *args: object, partial_result: object | None = None) -> None:
        self.partial_result = partial_result
        super().__init__(*args)


class TreeDirectoryCreationError(TreeUploadError):
    """Use quando um nivel da arvore nao puder ser resolvido ou criado."""

    pass


class TreeFileUploadError(TreeUploadError):
    """Use quando um arquivo falhar durante o upload sequencial da arvore."""

    pass


class TreeUploadCancelledError(TreeUploadError):
    """Use quando o chamador solicitar o cancelamento cooperativo da arvore."""


class SmallFileUploadError(UploadError):
    """Use quando falhar o fluxo de upload simples de arquivo pequeno,
    normalmente feito com PUT direto no endpoint de content."""

    def __init__(self, *args: object) -> None:
        super().__init__(*args)


class LargeFileUploadError(UploadError):
    """Use como erro generico do fluxo de upload grande, quando a falha
    pertence ao processo de envio resumivel mas nao a uma etapa mais
    especifica."""

    def __init__(self, *args: object) -> None:
        super().__init__(*args)


class LargeFileUploadNotSupportedError(UploadError):
    """Use quando um arquivo exceder o limite aceito pelo fluxo de upload
    pequeno e o fluxo atual nao puder ou nao deva continuar com upload
    grande."""

    def __init__(self, *args: object) -> None:
        super().__init__(*args)


class UploadSessionCreationError(UploadError):
    """Use quando a sessao de upload grande nao puder ser criada no Graph antes
    do envio dos chunks."""

    def __init__(self, *args: object) -> None:
        super().__init__(*args)


class UploadChunkError(UploadError):
    """Use quando um chunk do upload grande falhar durante transmissao,
    confirmacao ou resposta intermediaria do Graph."""

    def __init__(self, *args: object) -> None:
        super().__init__(*args)
