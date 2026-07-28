"""Laboratorio manual para validar o Core contra o Microsoft Graph.

Este arquivo nao faz parte da API do Core. Ele serve como roteiro executavel
para testar manualmente credenciais, resolucao de site, navegacao por drive,
leitura da raiz e iteracao paginada enquanto a implementacao evolui.
"""

import asyncio
from os import getenv

from dotenv import load_dotenv
from rich import print
from rich.traceback import install

from core import (
    GraphClientManager,
    GraphCredentials,
    SharePointService,
)

load_dotenv()
install(show_locals=True)


async def main() -> None:
    # Credenciais da aplicacao carregadas do ambiente local para autenticacao
    # no Microsoft Graph durante os testes manuais.
    client_id = getenv("CLIENT_ID", "")
    client_secret = getenv("CLIENT_SECRET", "")
    tenant_id = getenv("CLIENT_TENANT", "")

    # Instanciacao dos componentes principais do Core.
    credentials = GraphCredentials(
        client_id,
        client_secret,
        tenant_id,
    )

    graph_client_manager = GraphClientManager(credentials)
    sharepoint = SharePointService(graph_client_manager)

    site_url = "https://plangeconcombr.sharepoint.com/sites/RHConecta"

    # 1. Resolve a URL humana do SharePoint para um objeto de dominio.
    site = await sharepoint.resolve_site(site_url)

    print("[bold]Site resolvido:[/bold]")
    print(site)

    # 2. Lista todas as document libraries/drives disponiveis no site.
    drives = await sharepoint.list_site_drives(site)

    print("[bold]Drives encontrados:[/bold]")
    print(drives)

    # 3. Tenta localizar uma biblioteca especifica pelo nome.
    found_drive = await sharepoint.find_drive_by_name(
        name="RH & DP",
        site=site,
    )

    print("[bold]Drive encontrado por nome:[/bold]")
    print(found_drive)

    # 4. Recupera tambem a biblioteca padrao para funcionar como fallback.
    default_drive = await sharepoint.get_default_drive(site)

    print("[bold]Drive padrao:[/bold]")
    print(default_drive)

    # 5. Seleciona a biblioteca encontrada pelo nome ou, caso ela nao exista,
    # utiliza a biblioteca padrao do site.
    selected_drive = found_drive if found_drive is not None else default_drive

    print("[bold]Drive selecionado:[/bold]")
    print(selected_drive)

    # 6. Recupera o item raiz da biblioteca selecionada.
    root = await sharepoint.get_drive_root(selected_drive)

    print("[bold]Root do drive:[/bold]")
    print(root)

    # 7. Testa a operacao agregada que retorna os filhos da raiz.
    children = await sharepoint.list_children(
        selected_drive,
        root,
    )

    print("[bold]Resultado de list_children:[/bold]")
    print(children)

    # 8. Testa separadamente a iteracao paginada.
    #
    # Esta consulta e intencionalmente repetida para comparar o comportamento
    # de `list_children` com `iter_children`.
    print("[bold]Paginas retornadas por iter_children:[/bold]")

    async for page in sharepoint.iter_children(
        selected_drive,
        root,
    ):
        print(page)


if __name__ == "__main__":
    asyncio.run(main())
