"""System-level business invariants for NVL stock and open-PO."""

import unittest

from nvl.open_po import read_open_po, reconcile_target
from nvl.stock import reconcile_nvl_target
from tests.test_sync_nvl_open_po import (
    make_config as make_open_po_config,
    make_source_bytes as make_open_po_source,
    make_target_bytes as make_open_po_target,
)
from tests.test_sync_nvl_stock import make_mock_config, make_mock_target_bytes


class NVLBusinessContractTests(unittest.TestCase):
    def test_stock_missing_source_code_is_preserved_by_default(self):
        cfg = make_mock_config()
        target = make_mock_target_bytes(
            [
                ("VT001", 50),
                ("VT_MISSING", 777),
            ]
        )

        result = reconcile_nvl_target(target, {"VT001": 100.0}, cfg)

        by_code = {change["code"]: change for change in result.changes}
        self.assertEqual(by_code["VT001"]["after"], 100.0)
        self.assertNotIn("VT_MISSING", by_code)
        self.assertEqual(
            [item["code"] for item in result.missing_in_source],
            ["VT_MISSING"],
        )

    def test_stock_missing_source_can_be_fail_closed_by_policy(self):
        cfg = make_mock_config()
        cfg.preserve_missing_in_source = False
        target = make_mock_target_bytes(
            [
                ("VT001", 50),
                ("VT_MISSING", 777),
            ]
        )

        with self.assertRaisesRegex(RuntimeError, "preserve_missing_in_source"):
            reconcile_nvl_target(target, {"VT001": 100.0}, cfg)

    def test_open_po_remaining_quantity_is_clamped_per_line(self):
        totals, info = read_open_po(make_open_po_source(), make_open_po_config())

        self.assertEqual(totals, {"VT1": 80.0, "VT2": 0.0})
        self.assertEqual(info["over_received_rows"], 1)
        self.assertGreaterEqual(min(totals.values()), 0.0)

    def test_open_po_target_code_without_open_po_is_explicitly_zeroed(self):
        totals, _ = read_open_po(make_open_po_source(), make_open_po_config())
        _, expected = reconcile_target(
            make_open_po_target(),
            totals,
            make_open_po_config(),
        )

        self.assertEqual(expected["VT4"], 0.0)


if __name__ == "__main__":
    unittest.main()
