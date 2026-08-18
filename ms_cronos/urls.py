"""Helpers para validar URLs humanas do SharePoint e montar rotas Graph.

Este modulo lida apenas com strings de URL. Ele nao chama o Graph nem valida
se o site existe; essa responsabilidade pertence ao `SharePointService`.

Exemplo:
    graph_url = build_graph_site_url(
        'https://tenant.sharepoint.com/sites/RHConecta'
    )
"""

from pathlib import PurePosixPath
from urllib.parse import quote, urlparse

from core.errors import InvalidRemoteNameError, SharePointUrlError

INVALID_REMOTE_NAME_CHARACTERS = frozenset('"*:<>?/\\|')
MAX_REMOTE_NAME_LENGTH = 255
MAX_REMOTE_PATH_LENGTH = 400
RESERVED_REMOTE_NAMES = frozenset(
    {
        ".lock",
        "aux",
        "con",
        "desktop.ini",
        "nul",
        "prn",
        *(f"com{index}" for index in range(10)),
        *(f"lpt{index}" for index in range(10)),
    }
)


def validate_remote_name(name: str) -> str:
    """Valida um unico nome de arquivo ou pasta do SharePoint.

    O nome permanece decodificado. Percent-encoding pertence exclusivamente a
    fronteira que monta URLs para o Microsoft Graph.
    """
    if not name or name != name.strip():
        raise InvalidRemoteNameError(
            "O nome remoto nao pode ser vazio nem conter espacos nas extremidades: "
            f"{name!r}."
        )
    if len(name) > MAX_REMOTE_NAME_LENGTH:
        raise InvalidRemoteNameError(
            f"O nome remoto excede {MAX_REMOTE_NAME_LENGTH} caracteres: {name!r}."
        )
    if name in {".", ".."}:
        raise InvalidRemoteNameError(
            f"O segmento remoto {name!r} nao pode representar navegacao de caminho."
        )
    if name.endswith("."):
        raise InvalidRemoteNameError(
            f"O nome remoto nao pode terminar com ponto: {name!r}."
        )
    if any(character in INVALID_REMOTE_NAME_CHARACTERS for character in name):
        raise InvalidRemoteNameError(
            f"O nome remoto contem caracteres proibidos: {name!r}."
        )

    normalized_name = name.casefold()
    device_name = normalized_name.split(".", maxsplit=1)[0]
    if (
        normalized_name in RESERVED_REMOTE_NAMES
        or device_name in RESERVED_REMOTE_NAMES
        or normalized_name.startswith("~$")
        or "_vti_" in normalized_name
    ):
        raise InvalidRemoteNameError(
            f"O nome remoto e reservado pelo SharePoint: {name!r}."
        )
    return name


def validate_remote_path(path: PurePosixPath) -> PurePosixPath:
    """Valida um fragmento remoto relativo e todos os seus segmentos."""
    if path.is_absolute():
        raise InvalidRemoteNameError(
            f"O caminho remoto deve ser relativo ao item pai: '{path}'."
        )
    if path != PurePosixPath("."):
        for segment in path.parts:
            validate_remote_name(segment)
    if len(path.as_posix()) > MAX_REMOTE_PATH_LENGTH:
        raise InvalidRemoteNameError(
            "O fragmento remoto excede o limite de "
            f"{MAX_REMOTE_PATH_LENGTH} caracteres: '{path}'."
        )
    return path


def validate_graph_url(url: str, strict_validate: bool = False) -> bool:
    """Valida a forma minima de uma URL usada pelo Core.

    No modo padrao, valida apenas se a URL:
    - nao e vazia;
    - usa https;
    - possui hostname.

    No modo estrito, tambem exige um path explicito e rejeita caminhos que
    aparentam ja apontar para sub-recursos Graph.
    """
    if not url.strip():
        return False

    parts = urlparse(url)

    if parts.scheme != "https":
        return False

    if not parts.hostname:
        return False

    if strict_validate:
        if not parts.path:
            return False
        paths_fragments = parts.path.split(":")
        # Paths com mais de um fragmento separado por ':' geralmente indicam
        # que a URL ja aponta para um sub-recurso Graph.
        if len(paths_fragments) > 1:
            return False

        return True

    return True


def build_graph_site_url(sharepoint_url: str, strict_validate: bool = False) -> str:
    """Converte uma URL humana do SharePoint na URL Graph usada para resolver
    um site.

    Exemplo:
        https://tenant.sharepoint.com/sites/RHConecta
        -> https://graph.microsoft.com/v1.0/sites/
           tenant.sharepoint.com:/sites/RHConecta

    Quando a validacao falha, a funcao levanta SharePointUrlError para que a
    camada de servico nao precise conhecer detalhes do parse.
    """
    if not validate_graph_url(sharepoint_url, strict_validate):
        raise SharePointUrlError(
            "A URL do SharePoint deve ser uma URL HTTPS valida com hostname: "
            f"{sharepoint_url!r}."
        )
    parts = urlparse(sharepoint_url)

    # Sem path, a URL aponta para o site raiz do tenant e o Graph aceita apenas
    # o hostname como alvo de resolucao.
    if not parts.path:
        return f"https://graph.microsoft.com/v1.0/sites/{parts.hostname}".rstrip("/")
    # Com path, o Graph exige a sintaxe `hostname:/server-relative-path`.
    return f"https://graph.microsoft.com/v1.0/sites/{parts.hostname}:{
        parts.path
    }".rstrip("/")


def build_create_content_url(filename: str) -> str:
    """Monta o fragmento Graph usado para criar um arquivo por nome.

    O retorno e apenas o sufixo relativo ao item pai, por exemplo
    `:/curriculo.pdf:/content`. O nome permanece semantico nas camadas
    anteriores e recebe percent-encoding somente nesta fronteira de URL.
    """
    validate_remote_name(filename)
    encoded_filename = quote(filename, safe="")
    return f":/{encoded_filename}:/content"


def build_drive_create_content_url(
    drive_id: str,
    parent_item_id: str,
    target_path: str,
) -> str:
    """Monta a URL Graph completa usada para criar um arquivo pequeno.

    `target_path` deve conter apenas o fragmento relativo ao item pai, por
    exemplo `:/curriculo.pdf:/content`.
    """
    return (
        f"https://graph.microsoft.com/v1.0/drives/{drive_id}"
        f"/items/{parent_item_id}{target_path}"
    )
