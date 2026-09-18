"""Microsoft Graph authentication for local runs and GitHub Actions.

Modes:
- ``oidc``: GitHub Actions OIDC -> Microsoft Entra workload identity federation
  (production default in repository workflows).
- ``secret``: legacy MS_CLIENT_SECRET flow kept only for backward-compatible
  local tooling or external workloads that have not migrated yet.
- ``auto``: use OIDC when GitHub exposes its OIDC request variables, otherwise secret.

OIDC never silently falls back to a client secret. A broken trust configuration
must fail visibly instead of masking the problem with legacy credentials.
"""

from __future__ import annotations

import json
import os
from urllib.parse import urlsplit, urlunsplit, parse_qsl, urlencode

import msal
import requests


GRAPH_SCOPE = "https://graph.microsoft.com/.default"
DEFAULT_OIDC_AUDIENCE = "api://AzureADTokenExchange"
CLIENT_ASSERTION_TYPE = "urn:ietf:params:oauth:client-assertion-type:jwt-bearer"
VALID_AUTH_MODES = {"secret", "oidc", "auto"}


def _required_env(name: str) -> str:
    value = str(os.environ.get(name, "")).strip()
    if not value:
        raise RuntimeError(f"Thiếu biến môi trường bắt buộc: {name}")
    return value


def _legacy_secret_token() -> str:
    client_id = _required_env("MS_CLIENT_ID")
    tenant_id = _required_env("MS_TENANT_ID")
    client_secret = _required_env("MS_CLIENT_SECRET")

    app = msal.ConfidentialClientApplication(
        client_id=client_id,
        authority=f"https://login.microsoftonline.com/{tenant_id}",
        client_credential=client_secret,
    )
    result = app.acquire_token_for_client(scopes=[GRAPH_SCOPE])
    token = result.get("access_token")
    if not token:
        raise RuntimeError(
            "Không lấy được Microsoft Graph access token bằng Client Secret: "
            + json.dumps(result, ensure_ascii=False)
        )
    return token


def _append_query(url: str, **params: str) -> str:
    parts = urlsplit(url)
    query = dict(parse_qsl(parts.query, keep_blank_values=True))
    query.update(params)
    return urlunsplit((parts.scheme, parts.netloc, parts.path, urlencode(query), parts.fragment))


def request_github_oidc_token(
    *,
    audience: str = DEFAULT_OIDC_AUDIENCE,
    timeout: int = 30,
) -> str:
    """Request a short-lived GitHub Actions OIDC JWT for the Entra audience."""
    request_url = _required_env("ACTIONS_ID_TOKEN_REQUEST_URL")
    request_token = _required_env("ACTIONS_ID_TOKEN_REQUEST_TOKEN")

    response = requests.get(
        _append_query(request_url, audience=audience),
        headers={"Authorization": f"Bearer {request_token}"},
        timeout=timeout,
    )
    response.raise_for_status()
    payload = response.json()
    token = str(payload.get("value", "")).strip()
    if not token:
        raise RuntimeError("GitHub OIDC endpoint không trả về trường 'value'.")
    return token


def exchange_oidc_for_graph_token(
    oidc_token: str,
    *,
    tenant_id: str | None = None,
    client_id: str | None = None,
    timeout: int = 30,
) -> str:
    """Exchange the GitHub JWT for a Microsoft Graph access token via Entra."""
    tenant_id = str(tenant_id or _required_env("MS_TENANT_ID")).strip()
    client_id = str(client_id or _required_env("MS_CLIENT_ID")).strip()
    if not str(oidc_token or "").strip():
        raise RuntimeError("OIDC client assertion rỗng.")

    token_url = f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token"
    response = requests.post(
        token_url,
        data={
            "client_id": client_id,
            "scope": GRAPH_SCOPE,
            "grant_type": "client_credentials",
            "client_assertion_type": CLIENT_ASSERTION_TYPE,
            "client_assertion": oidc_token,
        },
        timeout=timeout,
    )
    if not response.ok:
        try:
            detail = response.json()
        except Exception:
            detail = response.text
        raise RuntimeError(
            f"Microsoft Entra từ chối GitHub OIDC ({response.status_code}): {detail}"
        )

    payload = response.json()
    token = str(payload.get("access_token", "")).strip()
    if not token:
        raise RuntimeError(
            "Microsoft Entra không trả access_token sau OIDC exchange: "
            + json.dumps(payload, ensure_ascii=False)
        )
    return token


def _github_oidc_available() -> bool:
    return bool(
        str(os.environ.get("ACTIONS_ID_TOKEN_REQUEST_URL", "")).strip()
        and str(os.environ.get("ACTIONS_ID_TOKEN_REQUEST_TOKEN", "")).strip()
    )


def get_access_token(mode: str | None = None) -> str:
    """Return a Graph access token using the explicitly selected auth mode.

    Repository production workflows explicitly default to ``oidc``. The function
    keeps ``secret`` as its library-level default only for backward-compatible
    local callers that do not set MS_AUTH_MODE. OIDC intentionally does not fall
    back to a secret when federation fails.
    """
    selected = str(mode or os.environ.get("MS_AUTH_MODE", "secret")).strip().casefold()
    if selected not in VALID_AUTH_MODES:
        raise RuntimeError(
            f"MS_AUTH_MODE={selected!r} không hợp lệ; dùng secret, oidc hoặc auto."
        )

    if selected == "auto":
        selected = "oidc" if _github_oidc_available() else "secret"

    if selected == "secret":
        return _legacy_secret_token()

    audience = str(os.environ.get("MS_OIDC_AUDIENCE", DEFAULT_OIDC_AUDIENCE)).strip()
    oidc_token = request_github_oidc_token(audience=audience)
    return exchange_oidc_for_graph_token(oidc_token)


__all__ = [
    "GRAPH_SCOPE",
    "DEFAULT_OIDC_AUDIENCE",
    "CLIENT_ASSERTION_TYPE",
    "get_access_token",
    "request_github_oidc_token",
    "exchange_oidc_for_graph_token",
]
