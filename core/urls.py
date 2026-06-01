from urllib.parse import urlparse


# INTERNAL PACKAGE
from core import (
    SharePointUrlError,
    )


def validate_graph_url(url: str, strict_validate: bool = False) -> bool:
    """
    Valida a forma minima de uma URL usada pelo Core.

    No modo padrao, valida apenas se a URL:
    - nao e vazia;
    - usa https;
    - possui hostname.

    No modo estrito, tambem exige um path explicito e rejeita caminhos que
    aparentam ja apontar para sub-recursos, como quando a URL contem mais de um
    fragmento separado por ':' no path.
    """
    # Se for vazia, não é uma URL válida
    if not url.strip():
        return False
    
    parts = urlparse(url)
    
    # Se não houver schema seguro é inválida
    if parts.scheme != 'https':
        return False
    
    # Se não houver host é inválida
    if not parts.hostname:
        return False
    
    if strict_validate:
        # Verifica se paths é vazio
        if not parts.path:
            return False
        paths_fragments = parts.path.split(':')
        # Verifica se o path tem mais de 1 fragmento
        if len(paths_fragments) > 1:
            # EX: GET https://graph.microsoft.com/v1.0/sites/{hostname}:/{site-server-relative-url} = Aceito, fragmento de Path = 1
            # EX: GET https://graph.microsoft.com/v1.0/sites/{hostname}:/{site-server-relative-url}:/lists/{list-id}/items/{item-id} Não aceito, fragmento de site = 2
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
    # Retorna uma URL formatada, se não, lança um erro

    # Verifica se a URL é valida (tem protocolo, não é vazia e tem host)
    if not validate_graph_url(sharepoint_url, strict_validate):
        raise SharePointUrlError
    parts = urlparse(sharepoint_url)
    
    # Se não houver complemento, retorna o site raiz
    if not parts.path:
        return f'https://graph.microsoft.com/v1.0/sites/{parts.hostname}'.rstrip('/')
    # Se não, retorna o caminho completo com o complemento para o recurso desejado
    return f'https://graph.microsoft.com/v1.0/sites/{parts.hostname}:{parts.path}'.rstrip('/')
