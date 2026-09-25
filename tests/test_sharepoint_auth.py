import os
import unittest
from unittest.mock import patch

import sharepoint.auth as auth


class FakeResponse:
    def __init__(self, *, status_code=200, payload=None, text=""):
        self.status_code = status_code
        self._payload = payload or {}
        self.text = text
        self.ok = 200 <= status_code < 300

    def json(self):
        return self._payload

    def raise_for_status(self):
        if not self.ok:
            raise RuntimeError(f"HTTP {self.status_code}")


class SharePointAuthTests(unittest.TestCase):
    def test_oidc_requests_github_token_with_azure_audience(self):
        env = {
            "ACTIONS_ID_TOKEN_REQUEST_URL": "https://oidc.example/token?api-version=2.0",
            "ACTIONS_ID_TOKEN_REQUEST_TOKEN": "runner-token",
        }
        with (
            patch.dict(os.environ, env, clear=False),
            patch("sharepoint.auth.requests.get", return_value=FakeResponse(payload={"value": "github-jwt"})) as get,
        ):
            token = auth.request_github_oidc_token()

        self.assertEqual(token, "github-jwt")
        called_url = get.call_args.args[0]
        self.assertIn("audience=api%3A%2F%2FAzureADTokenExchange", called_url)
        self.assertEqual(
            get.call_args.kwargs["headers"]["Authorization"],
            "Bearer runner-token",
        )

    def test_oidc_exchange_uses_client_assertion_flow(self):
        response = FakeResponse(payload={"access_token": "graph-token"})
        with patch("sharepoint.auth.requests.post", return_value=response) as post:
            token = auth.exchange_oidc_for_graph_token(
                "github-jwt",
                tenant_id="tenant-1",
                client_id="client-1",
            )

        self.assertEqual(token, "graph-token")
        data = post.call_args.kwargs["data"]
        self.assertEqual(data["client_id"], "client-1")
        self.assertEqual(data["scope"], auth.GRAPH_SCOPE)
        self.assertEqual(data["grant_type"], "client_credentials")
        self.assertEqual(data["client_assertion"], "github-jwt")
        self.assertEqual(data["client_assertion_type"], auth.CLIENT_ASSERTION_TYPE)

    def test_oidc_mode_does_not_fallback_to_secret(self):
        env = {
            "MS_AUTH_MODE": "oidc",
            "MS_TENANT_ID": "tenant",
            "MS_CLIENT_ID": "client",
        }
        with (
            patch.dict(os.environ, env, clear=True),
            patch("sharepoint.auth._legacy_secret_token") as legacy,
        ):
            with self.assertRaisesRegex(RuntimeError, "ACTIONS_ID_TOKEN_REQUEST_URL"):
                auth.get_access_token()
        legacy.assert_not_called()

    def test_auto_mode_uses_secret_outside_github_actions(self):
        with (
            patch.dict(os.environ, {"MS_AUTH_MODE": "auto"}, clear=True),
            patch("sharepoint.auth._legacy_secret_token", return_value="legacy") as legacy,
        ):
            self.assertEqual(auth.get_access_token(), "legacy")
        legacy.assert_called_once()

    def test_sharepoint_client_uses_canonical_auth_provider(self):
        import sharepoint.client as client
        self.assertIs(client.get_access_token, auth.get_access_token)



if __name__ == "__main__":
    unittest.main()
