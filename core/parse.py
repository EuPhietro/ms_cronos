"""Parsers semanticos entre o SDK do Microsoft Graph e os models do Core.

Este modulo concentra adaptadores puros: ele recebe models crus retornados pelo
SDK e devolve contratos internos, mais enxutos e previsiveis para o restante do
Core.

Responsabilidades principais:
- converter `Site`, `Drive` e `DriveItem` em referencias internas;
- converter envelopes `...CollectionResponse` em colecoes concretas do Core;
- adaptar caminhos locais em `LocalFile`;
- traduzir `ODataError` para a hierarquia de erros interna.

Exemplo:
    site_ref = parse_site(site)
    drives = parse_drive_collection_response(drive_response)
    local_file = parse_local_file(Path("/tmp/curriculo.pdf"))
"""

from __future__ import annotations
from pathlib import Path
from typing import Any

from msgraph.generated.models.drive import Drive
from msgraph.generated.models.drive_collection_response import DriveCollectionResponse
from msgraph.generated.models.drive_item import DriveItem
from msgraph.generated.models.drive_item_collection_response import DriveItemCollectionResponse
from msgraph.generated.models.o_data_errors.o_data_error import ODataError
from msgraph.generated.models.site import Site
from msgraph.generated.models.site_collection_response import SiteCollectionResponse

from core.errors import (
    DriveItemNotFoundError,
    DriveNotFoundError,
    FolderNotFoundError,
    GraphAuthenticationError,
    GraphPermissionError,
    GraphRequestError,
    GraphResourceConflictError,
    GraphResponseError,
    MSCronosError,
    NotAFileError,
    SiteResolutionError,
)
from core.models import (
    DriveItemCollection,
    DriveItemRef,
    DriveRef,
    DriveRefCollection,
    LocalFile,
    SiteRef,
    SiteRefCollection,
)


def parse_site(site: Site) -> SiteRef:
    """Converte um `Site` cru do SDK em `SiteRef`."""
    # O parser unitario adapta o model cru do SDK para a referencia enxuta
    # usada pelo Core.
    if not site.id:
        raise GraphResponseError("A resposta do Graph nao trouxe um site resolvido.")

    return SiteRef(
        id=site.id,
        name=site.name,
        display_name=site.display_name,
        web_url=site.web_url,
    )

def adapt_site(additional_data:dict[str, Any]) -> SiteRef:
    id = additional_data.get('id')
    name = additional_data.get('name')
    display_name = additional_data.get('displayName')
    web_url= additional_data.get('webUrl')
    if not id:
        raise GraphResponseError("A resposta do Graph nao trouxe um site resolvido.")
    return SiteRef(
        id,
        name,
        display_name,
        web_url
    )
    
def parse_drive(drive: Drive) -> DriveRef:
    """Converte um `Drive` cru do SDK em `DriveRef`."""
    # Drives do SDK carregam muitos detalhes; aqui o Core retira apenas o
    # contrato minimo que sera reutilizado nas proximas operacoes.
    if not drive.id:
        raise GraphResponseError("A resposta do Graph nao trouxe um drive valido.")

    return DriveRef(
        id=drive.id,
        name=drive.name,
        web_url=drive.web_url,
        drive_type=drive.drive_type,
    )


def parse_drive_item(drive_item: DriveItem) -> DriveItemRef:
    """Converte um `DriveItem` cru do SDK em `DriveItemRef`."""
    # A presenca das facets `folder` e `file` e o que permite distinguir com
    # seguranca se o item representa uma pasta ou um arquivo.
    if not drive_item.id:
        raise GraphResponseError("A resposta do Graph nao trouxe um item de drive valido.")

    return DriveItemRef(
        id=drive_item.id,
        name=drive_item.name,
        web_url=drive_item.web_url,
        is_folder=drive_item.folder is not None,
        is_file=drive_item.file is not None,
        size=drive_item.size,
    )


def parse_local_file(path: Path) -> LocalFile:
    """Converte um caminho local valido em `LocalFile`."""
    # Este parser adapta um caminho local para o modelo interno que depois pode
    # ser usado por upload e validacoes de arquivo.
    if not path.is_file():
        raise NotAFileError(f"O caminho informado nao aponta para um arquivo: {path}")

    return LocalFile(
        path=path,
        name=path.name,
        extension=path.suffix or None,
        size=path.stat().st_size,
    )


def parse_site_collection_response(
    site_collection_response: SiteCollectionResponse,
) -> SiteRefCollection:
    """Converte um envelope de sites do SDK em `SiteRefCollection`."""
    # Os envelopes de colecao do SDK chegam com `value`; o Core transforma cada
    # item e no final materializa a colecao semantica concreta.
    if not site_collection_response.value:
        raise GraphResponseError("A resposta do Graph nao trouxe sites resolvidos.")

    site_ref_collection = [parse_site(site) for site in site_collection_response.value]
    return SiteRefCollection.from_collection(site_ref_collection)


def parse_drive_collection_response(
    drive_collection_response: DriveCollectionResponse,
) -> DriveRefCollection:
    """Converte um envelope de drives do SDK em `DriveRefCollection`."""
    # O fluxo aqui e o mesmo de sites: validar o envelope, adaptar cada drive e
    # reconstruir a colecao concreta do Core.
    if not drive_collection_response.value:
        raise GraphResponseError("A resposta do Graph nao trouxe drives resolvidos.")

    drive_ref_collection = [parse_drive(drive) for drive in drive_collection_response.value]
    return DriveRefCollection.from_collection(drive_ref_collection)


def parse_drive_item_collection_response(
    drive_item_collection_response: DriveItemCollectionResponse,
) -> DriveItemCollection:
    """Converte um envelope de itens do SDK em `DriveItemCollection`."""
    # Itens de drive podem ser arquivos ou pastas, entao a adaptacao de cada
    # elemento preserva essas duas informacoes no modelo interno.
    if not drive_item_collection_response.value:
        raise GraphResponseError("A resposta do Graph nao trouxe itens de drive.")

    drive_item_ref_collection = [
        parse_drive_item(drive_item)
        for drive_item in drive_item_collection_response.value
    ]
    return DriveItemCollection.from_collection(drive_item_ref_collection)



def parse_o_data_error(
    o_data_error: ODataError,
    *,
    operation: str | None = None,
) -> MSCronosError:
    """Traduz `ODataError` do SDK para um erro semantico do Core.

    O parse considera primeiro o `error.code` retornado pelo Graph e depois, nos
    casos de recurso ausente, refina a classe final usando o contexto da
    operacao que falhou.
    """
    error = o_data_error.error
    code = (error.code or "").strip() if error else ""
    normalized_code = code.casefold()
    message = (
        error.message
        if error and error.message
        else o_data_error.primary_message or "O Microsoft Graph retornou um erro."
    )
    operation_name = (operation or "").casefold()

    if not error:
        return GraphRequestError(message)

    if normalized_code in {"accessdenied", "authorizationrequestdenied"}:
        return GraphPermissionError(message)

    if normalized_code in {"unauthenticated", "invalidauthenticationtoken"}:
        return GraphAuthenticationError(message)

    if normalized_code in {"namealreadyexists", "conflict"}:
        return GraphResourceConflictError(message)

    if normalized_code in {"invalidrequest", "badrequest", "badargument", "notallowed"}:
        return GraphRequestError(message)

    if normalized_code in {"generalexception", "servicenotavailable", "timeout", "toomanyrequests"}:
        return GraphRequestError(message)

    if normalized_code in {"itemnotfound", "resourcenotfound"}:
        if "site" in operation_name or "resolve_site" in operation_name:
            return SiteResolutionError(message)
        if "folder" in operation_name:
            return FolderNotFoundError(message)
        if "drive_item" in operation_name or "children" in operation_name or "root" in operation_name:
            return DriveItemNotFoundError(message)
        if "drive" in operation_name:
            return DriveNotFoundError(message)
        return GraphRequestError(message)

    return GraphRequestError(message)


