"""Regression tests cho safe resolver nguồn PO từ exact SharePoint URL."""

import unittest
from unittest.mock import patch

from sync_nvl_open_po import OpenPOConfig
from sync_nvl_open_po_safe import _graph_share_id, resolve_source_item_exact_url
from sharepoint.client import GraphRequestError


SOURCE_GUID = "1836226D-FB9D-4F8A-B6D5-3508A975936C"
SOURCE_ITEM_ID = "SOURCE-ITEM-123"
SOURCE_URL = (
    "https://vikodacomvn.sharepoint.com/:x:/r/sites/Planning/_layouts/15/Doc.aspx?"
    "sourcedoc=%7B1836226D-FB9D-4F8A-B6D5-3508A975936C%7D&"
    "file=BCTheodoiDMHANG.xlsm&action=default&mobileredirect=true"
)


def make_config():
    return OpenPOConfig(
        source_name="BCTheodoiDMHANG.xlsm",
        source_sourcedoc=SOURCE_GUID,
        source_sheet="REPORT_DONMUAHANG",
        source_start_row=6,
        source_code_col=6,
        source_purchase_col=9,
        source_received_col=11,
        source_status_col=28,
        target_name="Kế hoạch mua hàng.xlsx",
        target_path="Tinh san xuat Mua hang 2027/Kế hoạch mua hàng.xlsx",
        target_sourcedoc="89D1BA7B-006E-4527-B879-ABF120309214",
        target_sheet="Ton_NVL",
        target_start_row=2,
        target_code_col=1,
        target_value_col=5,
    )


class FakeGraph:
    def __init__(self, *, failures=0, guid=SOURCE_GUID, name="BCTheodoiDMHANG.xlsm", drive_id="drive-1"):
        self.failures = failures
        self.guid = guid
        self.name = name
        self.drive_id = drive_id
        self.urls = []

    def get_json(self, url, params=None):
        self.urls.append(url)
        if self.failures > 0:
            self.failures -= 1
            raise GraphRequestError(
                "Microsoft Graph lỗi 500",
                status_code=500,
                error_code="generalException",
            )
        return {
            "id": SOURCE_ITEM_ID,
            "name": self.name,
            "eTag": '"etag-1"',
            "sharepointIds": {"listItemUniqueId": self.guid},
            "parentReference": {"driveId": self.drive_id},
        }


class SafeOpenPOResolverTests(unittest.TestCase):
    def test_share_id_uses_graph_u_prefix(self):
        share_id = _graph_share_id(SOURCE_URL)
        self.assertTrue(share_id.startswith("u!"))
        self.assertNotIn("=", share_id)

    @patch("sync_nvl_open_po_safe._configured_source_url", return_value=SOURCE_URL)
    def test_resolves_by_exact_sharepoint_url_without_search(self, _):
        graph = FakeGraph()
        item = resolve_source_item_exact_url(graph, "drive-1", make_config())
        self.assertEqual(item["id"], SOURCE_ITEM_ID)
        self.assertEqual(len(graph.urls), 1)
        self.assertIn("/shares/u!", graph.urls[0])
        self.assertIn("/driveItem", graph.urls[0])
        self.assertNotIn("/search", graph.urls[0])
        self.assertNotIn("/items/", graph.urls[0])

    @patch("sync_nvl_open_po_safe._configured_source_url", return_value=SOURCE_URL)
    def test_transient_graph_500_is_retried(self, _):
        graph = FakeGraph(failures=1)
        with patch("sharepoint.client.time.sleep") as mocked_sleep:
            item = resolve_source_item_exact_url(graph, "drive-1", make_config())
        self.assertEqual(item["id"], SOURCE_ITEM_ID)
        self.assertEqual(len(graph.urls), 2)
        mocked_sleep.assert_called_once()

    @patch("sync_nvl_open_po_safe._configured_source_url", return_value=SOURCE_URL)
    def test_wrong_sourcedoc_is_rejected(self, _):
        graph = FakeGraph(guid="5D3E08DB-83EC-43AD-8F53-83D56B1C0D53")
        with self.assertRaises(RuntimeError):
            resolve_source_item_exact_url(graph, "drive-1", make_config())

    @patch("sync_nvl_open_po_safe._configured_source_url", return_value=SOURCE_URL)
    def test_wrong_name_is_rejected(self, _):
        graph = FakeGraph(name="Kế hoạch mua hàng.xlsx")
        with self.assertRaises(RuntimeError):
            resolve_source_item_exact_url(graph, "drive-1", make_config())

    @patch("sync_nvl_open_po_safe._configured_source_url", return_value=SOURCE_URL)
    def test_other_drive_is_rejected_with_clear_error(self, _):
        graph = FakeGraph(drive_id="drive-OTHER")
        with self.assertRaisesRegex(RuntimeError, "document library khác"):
            resolve_source_item_exact_url(graph, "drive-1", make_config())


if __name__ == "__main__":
    unittest.main()
