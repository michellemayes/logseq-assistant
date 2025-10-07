from __future__ import annotations

import logging

import msal

from .config import (
    build_token_cache,
    delegated_scopes,
    get_auth_mode,
    get_required_env,
    persist_token_cache,
)
from .constants import GRAPH_APP_SCOPE


def acquire_graph_token() -> str:
    client_id = get_required_env("MS_CLIENT_ID")
    tenant_id = get_required_env("MS_TENANT_ID")
    authority = f"https://login.microsoftonline.com/{tenant_id}"

    auth_mode = get_auth_mode()
    logging.debug("Using Microsoft Graph auth mode: %s", auth_mode)

    if auth_mode == "client_credentials":
        client_secret = get_required_env("MS_CLIENT_SECRET")
        app = msal.ConfidentialClientApplication(
            client_id=client_id,
            client_credential=client_secret,
            authority=authority,
        )
        token_response = app.acquire_token_for_client(scopes=GRAPH_APP_SCOPE)
        if "access_token" not in token_response:
            raise RuntimeError(f"Failed to acquire Graph token: {token_response}")
        return token_response["access_token"]

    cache = build_token_cache()
    app = msal.PublicClientApplication(
        client_id=client_id,
        authority=authority,
        token_cache=cache,
    )
    scopes = delegated_scopes()
    accounts = app.get_accounts()
    result = None
    if accounts:
        logging.debug(
            "Attempting silent token acquisition for account %s",
            accounts[0].get("username"),
        )
        result = app.acquire_token_silent(scopes=scopes, account=accounts[0])

    if not result:
        logging.info("Initiating device code flow for delegated Graph access")
        flow = app.initiate_device_flow(scopes=scopes)
        if "user_code" not in flow:
            raise RuntimeError(f"Failed to initiate device flow: {flow}")
        logging.warning(flow["message"])
        result = app.acquire_token_by_device_flow(flow)

    if "access_token" not in result:
        error = result.get("error")
        description = result.get("error_description", "")
        if error == "invalid_client" and "AADSTS7000218" in description:
            raise RuntimeError(
                "Failed to acquire delegated token: AADSTS7000218. Enable "
                "'Allow public client flows' on the Azure AD app registration "
                "or set MS_CLIENT_SECRET and MS_AUTH_MODE=client_credentials to use "
                "application permissions."
            )
        raise RuntimeError(f"Failed to acquire delegated token: {result}")

    persist_token_cache(cache)
    return result["access_token"]
