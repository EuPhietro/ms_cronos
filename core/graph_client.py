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
        scopes: str | Sequence[str] | None = None,
    ) -> None:
        self._credentials = credentials
        if scopes is None:
            self._scopes = DEFAULT_SCOPES
        elif isinstance(scopes, str):
            self._scopes = (scopes,)
        else:
            self._scopes = tuple(scopes)

        self._validate_credentials(credentials)
        self._validate_scopes(self._scopes)

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
            raise RuntimeError(
                "O cliente do Microsoft Graph nao pode ser usado depois que "
                "o GraphClientManager foi fechado."
            )
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
        if not credentials.client_id.strip():
            raise ValueError("GraphCredentials.client_id nao pode ser vazio.")
        if not credentials.client_secret.strip():
            raise ValueError("GraphCredentials.client_secret nao pode ser vazio.")
        if not credentials.tenant_id.strip():
            raise ValueError("GraphCredentials.tenant_id nao pode ser vazio.")

    @staticmethod
    def _validate_scopes(scopes: Sequence[str]) -> None:
        if not scopes:
            raise ValueError("Informe ao menos um scope do Microsoft Graph.")
        for scope in scopes:
            if not isinstance(scope, str) or not scope.strip():
                raise ValueError(
                    "Cada scope do Microsoft Graph deve ser uma string nao vazia."
                )


def create_graph_client_manager(
    credentials: GraphCredentials,
    scopes: str | Sequence[str] | None = None,
) -> GraphClientManager:
    """Cria um `GraphClientManager` com os scopes informados ou os padroes."""
    return GraphClientManager(credentials=credentials, scopes=scopes)


def create_graph_client(
    credentials: GraphCredentials,
    scopes: str | Sequence[str] | None = None,
) -> GraphClientManager:
    """Alias de compatibilidade para criar o manager do Graph."""
    return create_graph_client_manager(credentials=credentials, scopes=scopes)
