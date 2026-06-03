"""Helpers para validar URLs humanas do SharePoint e montar rotas Graph.

Este modulo lida apenas com strings de URL. Ele nao chama o Graph e nao valida se
o site existe; essa responsabilidade pertence ao `SharePointService`.

Exemplo:
    graph_url = build_graph_site_url(
        'https://tenant.sharepoint.com/sites/RHConecta'
    )
"""

from urllib.parse import urlparse

from core.errors import SharePointUrlError


def validate_graph_url(url: str, strict_validate: bool = False) -> bool:
    """
    Valida a forma minima de uma URL usada pelo Core.

    No modo padrao, valida apenas se a URL:
    - nao e vazia;
    - usa https;
    - possui hostname.

    No modo estrito, tambem exige um path explicito e rejeita caminhos que
    aparentam ja apontar para sub-recursos Graph.
    """
    # URLs vazias ou compostas apenas por espacos nao podem ser resolvidas.
    if not url.strip():
        return False

    parts = urlparse(url)

    # O Core exige HTTPS para URLs humanas do SharePoint.
    if parts.scheme != 'https':
        return False

    # Sem hostname nao ha tenant SharePoint para resolver.
    if not parts.hostname:
        return False

    if strict_validate:
        # No modo estrito, exigimos path explicito de site.
        if not parts.path:
            return False
        paths_fragments = parts.path.split(':')
        # Paths com mais de um fragmento separado por ':' geralmente indicam
        # que a URL ja aponta para um sub-recurso Graph.
        if len(paths_fragments) > 1:
            return False

        return True

    return True


def build_graph_site_url(sharepoint_url: str, strict_validate: bool = False) -> str:
    """
    Converte uma URL humana do SharePoint na URL Graph usada para resolver um
    site.

    Exemplo:
        https://tenant.sharepoint.com/sites/RHConecta
        -> https://graph.microsoft.com/v1.0/sites/tenant.sharepoint.com:/sites/RHConecta

    Quando a validacao falha, a funcao levanta SharePointUrlError para que a
    camada de servico nao precise conhecer detalhes do parse.
    """
    # A validacao fica centralizada para manter o servico livre de detalhes de
    # parse de URL.
    if not validate_graph_url(sharepoint_url, strict_validate):
        raise SharePointUrlError
    parts = urlparse(sharepoint_url)

    # Sem path, a URL aponta para o site raiz do tenant.
    if not parts.path:
        return f'https://graph.microsoft.com/v1.0/sites/{parts.hostname}'.rstrip('/')
    # Com path, usamos a sintaxe Graph `hostname:/server-relative-path`.
    return f'https://graph.microsoft.com/v1.0/sites/{parts.hostname}:{parts.path}'.rstrip('/')


def build_create_content_url(filename: str) -> str:
    """Monta o fragmento Graph usado para criar um arquivo por nome.

    O retorno e apenas o sufixo relativo ao item pai, por exemplo
    `:/curriculo.pdf:/content`.
    """
    return f':/{filename}:/content'


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
        f'https://graph.microsoft.com/v1.0/drives/{drive_id}'
        f'/items/{parent_item_id}{target_path}'
    )
