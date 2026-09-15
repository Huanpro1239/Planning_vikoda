import unittest
from unittest.mock import patch

from sharepoint.client import (
    SharePointClient,
    GraphClient,
    GraphRequestError,
    assert_same_revision,
    graph_share_id,
    resolve_share_url,
    validate_item_identity,
)


GUID = "1836226D-FB9D-4F8A-B6D5-3508A975936C"
URL = "https://vikodacomvn.sharepoint.com/:x:/r/sites/Planning/_layouts/15/Doc.aspx?sourcedoc=%7B1836226D-FB9D-4F8A-B6D5-3508A975936C%7D&file=BCTheodoiDMHANG.xlsm"


class FakeGraph:
    def __init__(self, responses):
        self.responses = list(responses)
        self.urls = []

    def get_json(self, url, params=None):
        self.urls.append((url, params))
        value = self.responses.pop(0)
        if isinstance(value, Exception):
            raise value
        return value


class SharePointClientTests(unittest.TestCase):
    def test_sharepoint_client_is_single_graph_client_class(self):
        self.assertIs(SharePointClient, GraphClient)

    def test_graph_share_id_is_urlsafe_and_unpadded(self):
        token = graph_share_id(URL)
        self.assertTrue(token.startswith("u!"))
        self.assertNotIn("=", token)

    def test_validate_identity_flattens_remote_item(self):
        item = {
            "remoteItem": {
                "id": "item-1",
                "name": "BCTheodoiDMHANG.xlsm",
                "sharepointIds": {"listItemUniqueId": GUID.lower()},
                "parentReference": {"driveId": "drive-1"},
            }
        }
        result = validate_item_identity(
            item,
            expected_name="BCTheodoiDMHANG.xlsm",
            expected_sourcedoc=GUID,
            expected_drive_id="drive-1",
        )
        self.assertEqual(result["id"], "item-1")

    def test_resolve_share_url_retries_transient_500(self):
        graph = FakeGraph([
            GraphRequestError("500", status_code=500, error_code="generalException"),
            {
                "id": "item-1",
                "name": "BCTheodoiDMHANG.xlsm",
                "sharepointIds": {"listItemUniqueId": GUID},
                "parentReference": {"driveId": "drive-1"},
            },
        ])
        with patch("sharepoint.client.time.sleep") as sleep:
            result = resolve_share_url(
                graph,
                URL,
                expected_name="BCTheodoiDMHANG.xlsm",
                expected_sourcedoc=GUID,
                expected_drive_id="drive-1",
            )
        self.assertEqual(result["id"], "item-1")
        self.assertEqual(len(graph.urls), 2)
        sleep.assert_called_once()

    def test_revision_change_raises_precondition_failure(self):
        with self.assertRaises(GraphRequestError) as ctx:
            assert_same_revision({"eTag": "v1"}, {"eTag": "v2"}, label="source")
        self.assertEqual(ctx.exception.status_code, 412)
        self.assertEqual(ctx.exception.error_code, "preconditionFailed")


if __name__ == "__main__":
    unittest.main()
