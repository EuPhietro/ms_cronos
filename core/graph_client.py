"""Gerenciamento do cliente autenticado do Microsoft Graph.

Este modulo cria o `GraphServiceClient` a partir de credenciais ja resolvidas
por outra camada. Ele nao le `.env` e nao conhece regras de SharePoint; sua
responsabilidade e apenas construir, expor e fechar o client.

Exemplo:
    manager = GraphClientManager(credentials)
    client = manager.client
    await manager.close()
"""

from __future__ import annotations

from collections.abc import Sequence
from types import TracebackType

from azure.identity.aio import ClientSecretCredential
from msgraph.graph_service_client import GraphServiceClient

from core.models import GraphCredentials

DEFAULT_SCOPES: tuple[str, ...] = ("https://graph.microsoft.com/.default",)


class GraphClientManager:
    """Responsavel por criar e fechar o ciclo de vida do GraphServiceClient.

    Esta classe nao le variaveis de ambiente. Ela recebe um objeto de
    credenciais ja resolvido por outra camada e controla a vida util da
    credencial assincrona do Azure.
    """

    def __init__(
        self,
        credentials: GraphCredentials,
        scopes: Sequence[str] | None = None,
    ) -> None:
        # O manager preserva o contrato de entrada para que outras camadas
        # possam inspecionar credenciais e scopes sem recriar o client.
        self._credentials = credentials
        self._scopes = tuple(scopes or DEFAULT_SCOPES)

        # Falhas de configuracao devem acontecer cedo, antes da criacao da
        # credencial assincrona do Azure e do client Graph.
        self._validate_credentials(credentials)
        self._validate_scopes(self._scopes)

        # A credencial e o client ficam encapsulados no manager para permitir
        # fechamento controlado e uso com `async with`.
        self._credential = ClientSecretCredential(
            tenant_id=credentials.tenant_id,
            client_id=credentials.client_id,
            client_secret=credentials.client_secret,
        )
        self._client = GraphServiceClient(self._credential, scopes=list(self._scopes))
        self._closed = False

    @property
    def credential(self) -> ClientSecretCredential:
        """Retorna a credencial assíncrona usada pelo client Graph."""
        return self._credential

    @property
    def client(self) -> GraphServiceClient:
        """Retorna o `GraphServiceClient` enquanto o manager estiver aberto."""
        if self._closed:
            raise RuntimeError("Graph client manager is closed.")
        return self._client

    @property
    def scopes(self) -> tuple[str, ...]:
        """Retorna os scopes configurados para autenticacao no Graph."""
        return self._scopes

    async def close(self) -> None:
        """Fecha a credencial subjacente usada pelo client Graph."""
        if self._closed:
            return
        await self._credential.close()
        self._closed = True

    async def __aenter__(self) -> GraphClientManager:
        """Permite usar o manager com `async with`."""
        return self

    async def __aexit__(
        self,
        exc_type: type[BaseException] | None,
        exc: BaseException | None,
        tb: TracebackType | None,
    ) -> None:
        """Garante fechamento da credencial ao sair do contexto assincrono."""
        await self.close()

    @staticmethod
    def _validate_credentials(credentials: GraphCredentials) -> None:
        # O Core exige strings nao vazias para evitar erros obscuros mais tarde
        # na autenticacao remota.
        if not credentials.client_id.strip():
            raise ValueError("GraphCredentials.client_id cannot be empty.")
        if not credentials.client_secret.strip():
            raise ValueError("GraphCredentials.secrets_token cannot be empty.")
        if not credentials.tenant_id.strip():
            raise ValueError("GraphCredentials.tenant_id cannot be empty.")

    @staticmethod
    def _validate_scopes(scopes: Sequence[str]) -> None:
        # Scopes vazios quebrariam a construcao do client e precisam ser
        # rejeitados antes de qualquer chamada remota.
        if not scopes:
            raise ValueError("At least one scope must be provided.")
        for scope in scopes:
            if not isinstance(scope, str) or not scope.strip():
                raise ValueError("Scopes must be non-empty strings.")


def create_graph_client_manager(
    credentials: GraphCredentials,
    scopes: Sequence[str] | None = None,
) -> GraphClientManager:
    """Cria um `GraphClientManager` com os scopes informados ou os padroes."""
    return GraphClientManager(credentials=credentials, scopes=scopes)


def create_graph_client(
    credentials: GraphCredentials,
    scopes: Sequence[str] | None = None,
) -> GraphClientManager:
    """Alias de compatibilidade para criar o manager do Graph."""
    return create_graph_client_manager(credentials=credentials, scopes=scopes)
