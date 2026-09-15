"""Regression tests cho resolver PO nguồn đã pin bằng SharePoint item ID."""

import unittest
from unittest.mock import patch

from sync_nvl_open_po import OpenPOConfig
from sync_nvl_open_po_safe import resolve_source_item_pinned
from sync_stock import GraphRequestError


SOURCE_ITEM_ID = "01EX3DDZD3XLIYS3QAE5C3Q6NL6EQDBEQU"
SOURCE_GUID = "1836226D-FB9D-4F8A-B6D5-3508A975936C"


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
    def __init__(self, *, failures=0, guid=SOURCE_GUID):
        self.failures = failures
        self.guid = guid
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
            "name": "BCTheodoiDMHANG.xlsm",
            "eTag": '"etag-1"',
            "sharepointIds": {"listItemUniqueId": self.guid},
        }


class SafeOpenPOResolverTests(unittest.TestCase):
    def test_resolves_by_exact_item_id_without_search(self):
        graph = FakeGraph()
        item = resolve_source_item_pinned(
            graph,
            "drive-1",
            make_config(),
            item_id=SOURCE_ITEM_ID,
        )
        self.assertEqual(item["id"], SOURCE_ITEM_ID)
        self.assertEqual(len(graph.urls), 1)
        self.assertIn(f"/items/{SOURCE_ITEM_ID}", graph.urls[0])
        self.assertNotIn("/search", graph.urls[0])

    def test_transient_graph_500_is_retried(self):
        graph = FakeGraph(failures=1)
        with patch("sync_nvl_open_po_safe.time.sleep") as mocked_sleep:
            item = resolve_source_item_pinned(
                graph,
                "drive-1",
                make_config(),
                item_id=SOURCE_ITEM_ID,
            )
        self.assertEqual(item["id"], SOURCE_ITEM_ID)
        self.assertEqual(len(graph.urls), 2)
        mocked_sleep.assert_called_once()

    def test_wrong_sourcedoc_is_rejected(self):
        graph = FakeGraph(guid="5D3E08DB-83EC-43AD-8F53-83D56B1C0D53")
        with self.assertRaises(RuntimeError):
            resolve_source_item_pinned(
                graph,
                "drive-1",
                make_config(),
                item_id=SOURCE_ITEM_ID,
            )


if __name__ == "__main__":
    unittest.main()
