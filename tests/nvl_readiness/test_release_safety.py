"""NVL release-safety checks used by the production-readiness gate."""

from __future__ import annotations

from hashlib import sha256
from pathlib import Path
import tempfile
import unittest

from nvl.stock import run_nvl_sync
from tests.test_nvl_publish_boundary import FakeGraphClient
from tests.test_sync_nvl_stock import (
    make_mock_config,
    make_mock_source_bytes,
    make_mock_target_bytes,
)


class NVLReleaseSafetyTests(unittest.TestCase):
    def setUp(self):
        self.cfg = make_mock_config()
        self.cfg.source_path = "path/to/XNT_ketoan_Vikoda.xlsm"
        self.cfg.target_path = "path/to/Kế hoạch mua hàng.xlsx"
        self.source_bytes = make_mock_source_bytes(
            [("VT001", 100), ("VT002", 200)]
        )
        self.target_bytes = make_mock_target_bytes(
            [("VT001", 50), ("VT002", 70)]
        )

    def _offline_proposal(self, root: Path) -> bytes:
        source = root / "source.xlsm"
        target = root / "target.xlsx"
        out = root / "out"
        source.write_bytes(self.source_bytes)
        target.write_bytes(self.target_bytes)

        report = run_nvl_sync(
            self.cfg,
            source_file=source,
            target_file=target,
            out_dir=out,
            publish=False,
        )
        self.assertEqual(report["mode"], "offline")
        proposal = out / "nvl_stock_proposal.xlsx"
        self.assertTrue(proposal.is_file())
        return proposal.read_bytes()

    def test_same_inputs_produce_identical_proposal(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            root = Path(tmpdir)
            first = self._offline_proposal(root / "run1")
            second = self._offline_proposal(root / "run2")

        self.assertEqual(sha256(first).hexdigest(), sha256(second).hexdigest())
        self.assertEqual(first, second)

    def test_dry_run_produces_auditable_artifacts_without_upload(self):
        graph = FakeGraphClient()
        graph.set_file(
            self.cfg.source_path,
            self.source_bytes,
            etag="src-etag-1",
            item_id="src-1",
        )
        graph.set_file(
            self.cfg.target_path,
            self.target_bytes,
            etag="tgt-etag-1",
            item_id="tgt-1",
        )

        with tempfile.TemporaryDirectory() as tmpdir:
            report = run_nvl_sync(
                self.cfg,
                out_dir=tmpdir,
                publish=False,
                graph=graph,
            )
            out = Path(tmpdir)
            self.assertEqual(report["mode"], "dry_run")
            self.assertEqual(len(graph.uploads), 0)
            self.assertTrue((out / "nvl_stock_proposal.xlsx").is_file())
            self.assertTrue((out / "nvl_stock_report.json").is_file())


if __name__ == "__main__":
    unittest.main()
