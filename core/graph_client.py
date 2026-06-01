from __future__ import annotations

from collections.abc import Sequence

from azure.identity.aio import ClientSecretCredential
from msgraph.graph_service_client import GraphServiceClient

from core.models import GraphCredentials


DEFAULT_SCOPES: tuple[str, ...] = ("https://graph.microsoft.com/.default",)


class GraphClientManager:
    """
    Responsavel por criar e fechar o ciclo de vida do GraphServiceClient.

    Esta classe nao le variaveis de ambiente. Ela recebe um objeto de
    credenciais ja resolvido por outra camada.
    """

    def __init__(
        self,
        credentials: GraphCredentials, # Recebe por parâmetro o model Graph credentials do azure.identity.aio
        scopes: Sequence[str] | None = None, # Representa uma sequencia de escopo personalizado
    ) -> None:
        self._credentials = credentials 
        self._scopes = tuple(scopes or DEFAULT_SCOPES)

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
        '''Atributo calculado que retorna credentials'''
        return self._credential

    @property
    def client(self) -> GraphServiceClient:
        '''Atributo calculado que retorna _client'''
        if self._closed:
            raise RuntimeError("Graph client manager is closed.")
        return self._client

    @property
    def scopes(self) -> tuple[str, ...]:
        '''Atributo calculado que retorna _scope'''
        return self._scopes

    async def close(self) -> None:
        '''Função que fecha a conexão com o a API'''
        if self._closed:
            return
        await self._credential.close()
        self._closed = True

    # Usado pelo gerenciador de contexto With
    async def __aenter__(self) -> GraphClientManager:
        return self
# Usado pelo gerenciador de contexto With
    async def __aexit__(self, exc_type, exc, tb) -> None:
        await self.close()

    @staticmethod
    def _validate_credentials(credentials: GraphCredentials) -> None:
        if not credentials.client_id.strip():
            raise ValueError("GraphCredentials.client_id cannot be empty.")
        if not credentials.client_secret.strip():
            raise ValueError("GraphCredentials.secrets_token cannot be empty.")
        if not credentials.tenant_id.strip():
            raise ValueError("GraphCredentials.tenant_id cannot be empty.")

    @staticmethod
    def _validate_scopes(scopes: Sequence[str]) -> None:
        if not scopes:
            raise ValueError("At least one scope must be provided.")
        for scope in scopes:
            if not isinstance(scope, str) or not scope.strip():
                raise ValueError("Scopes must be non-empty strings.")


def create_graph_client_manager(
    credentials: GraphCredentials,
    scopes: Sequence[str] | None = None,
) -> GraphClientManager:
    return GraphClientManager(credentials=credentials, scopes=scopes)


def create_graph_client(
    credentials: GraphCredentials,
    scopes: Sequence[str] | None = None,
) -> GraphClientManager:
    return create_graph_client_manager(credentials=credentials, scopes=scopes)
