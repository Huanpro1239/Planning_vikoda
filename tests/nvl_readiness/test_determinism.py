"""Deterministic proposal-output contracts for NVL workbooks."""

import unittest

from nvl.open_po import (
    patch_target_workbook,
    read_open_po,
    reconcile_target,
    verify_target,
)
from nvl.stock import (
    patch_nvl_destination_workbook,
    reconcile_nvl_target,
    verify_nvl_patched_workbook,
)
from tests.test_sync_nvl_open_po import (
    make_config as make_open_po_config,
    make_source_bytes as make_open_po_source,
    make_target_bytes as make_open_po_target,
)
from tests.test_sync_nvl_stock import make_mock_config, make_mock_target_bytes


class NVLDeterminismTests(unittest.TestCase):
    def test_stock_proposal_bytes_are_deterministic(self):
        cfg = make_mock_config()
        target = make_mock_target_bytes(
            [
                ("VT001", 10),
                ("VT002", 20),
            ]
        )
        result = reconcile_nvl_target(
            target,
            {"VT001": 100.0, "VT002": 200.0},
            cfg,
        )

        first = patch_nvl_destination_workbook(target, result, cfg)
        second = patch_nvl_destination_workbook(target, result, cfg)

        self.assertEqual(first, second)
        self.assertTrue(
            verify_nvl_patched_workbook(target, first, result, cfg)["ok"]
        )

    def test_open_po_proposal_bytes_are_deterministic(self):
        cfg = make_open_po_config()
        source = make_open_po_source()
        target = make_open_po_target()
        totals, _ = read_open_po(source, cfg)
        changes, expected = reconcile_target(target, totals, cfg)

        first = patch_target_workbook(target, changes, cfg)
        second = patch_target_workbook(target, changes, cfg)

        self.assertEqual(first, second)
        verify_target(first, expected, cfg)


if __name__ == "__main__":
    unittest.main()
