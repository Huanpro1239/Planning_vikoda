"""Integration & boundary tests for sync_nvl_stock using mock Graph."""

from io import BytesIO
import json
import tempfile
from pathlib import Path
from typing import Any
import unittest

from openpyxl import Workbook

from sync_nvl_stock import NVLConfig, run_nvl_sync
from sync_stock import GraphRequestError
from tests.test_sync_nvl_stock import make_mock_config, make_mock_source_bytes, make_mock_target_bytes


class FakeGraphClient:
    def __init__(self, *, first_upload_412=False, first_upload_429=False):
        self.site_id = "fake-site-id"
        self.drive_id = "fake-drive-id"
        self.first_upload_412 = first_upload_412
        self.first_upload_429 = first_upload_429
        self.upload_attempts = 0
        self.uploads: list[dict[str, Any]] = []
        self.files: dict[str, bytes] = {}
        self.items: dict[str, dict[str, Any]] = {}
        self.target_fetch_count = 0

    def get_site_id(self):
        return self.site_id

    def get_default_drive_id(self, site_id):
        return self.drive_id

    def set_file(self, path: str, content: bytes, etag: str = "etag-1", item_id: str = "item-1"):
        self.files[item_id] = content
        self.items[path] = {
            "id": item_id,
            "name": Path(path).name,
            "eTag": etag,
            "size": len(content),
            "lastModifiedDateTime": "2026-09-09T00:00:00Z",
        }

    def get_item_by_path(self, drive_id, path):
        if path not in self.items:
            raise GraphRequestError(f"File not found: {path}", status_code=404, error_code="itemNotFound")
        item = dict(self.items[path])
        if "Kế hoạch mua hàng" in path:
            self.target_fetch_count += 1
            if self.first_upload_412 and self.target_fetch_count > 1:
                item["eTag"] = "etag-updated-after-412"
        return item

    def download_file(self, drive_id, item_id):
        if item_id not in self.files:
            raise GraphRequestError(f"Item not found: {item_id}", status_code=404, error_code="itemNotFound")
        return self.files[item_id]

    def upload_file(self, drive_id, item_id, content, expected_etag):
        self.upload_attempts += 1
        self.uploads.append({
            "drive_id": drive_id,
            "item_id": item_id,
            "content": content,
            "expected_etag": expected_etag,
        })

        if self.first_upload_429 and self.upload_attempts == 1:
            err = GraphRequestError("Too Many Requests", status_code=429, error_code="tooManyRequests")
            err.retry_after_seconds = 0.05
            raise err

        if self.first_upload_412 and self.upload_attempts == 1:
            raise GraphRequestError(
                "preconditionFailed: file changed",
                status_code=412,
                error_code="preconditionFailed",
            )

        return {"id": item_id, "name": "uploaded.xlsx", "eTag": "etag-new"}


class NVLPublishBoundaryTests(unittest.TestCase):
    def setUp(self):
        self.cfg = make_mock_config()
        self.cfg.source_path = "path/to/XNT_ketoan_Vikoda.xlsm"
        self.cfg.target_path = "path/to/Kế hoạch mua hàng.xlsx"

        self.source_bytes = make_mock_source_bytes([
            ("VT001", 100),
            ("VT002", 200),
        ])
        self.target_bytes = make_mock_target_bytes([
            ("VT001", 50),
            ("VT002", 70),
        ])

    def test_dry_run_does_not_upload(self):
        """Chế độ dry-run (publish=False) đọc snapshot, tạo proposal/report, nhưng không gọi upload."""
        fake_graph = FakeGraphClient()
        fake_graph.set_file(self.cfg.source_path, self.source_bytes, etag="src-etag-1", item_id="src-1")
        fake_graph.set_file(self.cfg.target_path, self.target_bytes, etag="tgt-etag-1", item_id="tgt-1")

        with tempfile.TemporaryDirectory() as tmpdir:
            report = run_nvl_sync(self.cfg, out_dir=tmpdir, publish=False, graph=fake_graph)

            self.assertEqual(report["mode"], "dry_run")
            self.assertEqual(report["metrics"]["changed_count"], 2)
            self.assertEqual(len(fake_graph.uploads), 0, "Dry-run tuyệt đối không được gọi upload_file!")

            # Files proposal và report phải được tạo
            self.assertTrue((Path(tmpdir) / "nvl_stock_proposal.xlsx").exists())
            self.assertTrue((Path(tmpdir) / "nvl_stock_report.json").exists())

    def test_publish_with_changes_calls_upload_with_expected_etag(self):
        """Chế độ publish=True khi có thay đổi gọi upload_file với đúng expected_etag."""
        fake_graph = FakeGraphClient()
        fake_graph.set_file(self.cfg.source_path, self.source_bytes, etag="src-etag-1", item_id="src-1")
        fake_graph.set_file(self.cfg.target_path, self.target_bytes, etag="tgt-etag-1", item_id="tgt-1")

        with tempfile.TemporaryDirectory() as tmpdir:
            report = run_nvl_sync(self.cfg, out_dir=tmpdir, publish=True, graph=fake_graph)

            self.assertEqual(report["mode"], "publish")
            self.assertEqual(report["metrics"]["changed_count"], 2)
            self.assertEqual(len(fake_graph.uploads), 1)
            self.assertEqual(fake_graph.uploads[0]["expected_etag"], "tgt-etag-1")
            self.assertIn("upload_result", report)

    def test_publish_without_changes_skips_upload(self):
        """Chế độ publish=True khi dữ liệu đã khớp hoàn toàn (0 thay đổi) sẽ bỏ qua upload."""
        # Dữ liệu đích đã có sẵn VT001=100, VT002=200
        identical_target = make_mock_target_bytes([
            ("VT001", 100),
            ("VT002", 200),
        ])
        fake_graph = FakeGraphClient()
        fake_graph.set_file(self.cfg.source_path, self.source_bytes, etag="src-etag-1", item_id="src-1")
        fake_graph.set_file(self.cfg.target_path, identical_target, etag="tgt-etag-1", item_id="tgt-1")

        with tempfile.TemporaryDirectory() as tmpdir:
            report = run_nvl_sync(self.cfg, out_dir=tmpdir, publish=True, graph=fake_graph)

            self.assertEqual(report["status"], "unchanged")
            self.assertEqual(report["metrics"]["changed_count"], 0)
            self.assertEqual(len(fake_graph.uploads), 0, "Không có thay đổi thì không gọi upload!")

    def test_http_412_refetches_and_retries_upload_with_new_etag(self):
        """Xử lý xung đột HTTP 412: tải lại đích, tính lại và upload với ETag mới."""
        fake_graph = FakeGraphClient(first_upload_412=True)
        fake_graph.set_file(self.cfg.source_path, self.source_bytes, etag="src-etag-1", item_id="src-1")
        fake_graph.set_file(self.cfg.target_path, self.target_bytes, etag="tgt-etag-1", item_id="tgt-1")

        with tempfile.TemporaryDirectory() as tmpdir:
            report = run_nvl_sync(self.cfg, out_dir=tmpdir, publish=True, graph=fake_graph, max_publish_attempts=2)

            self.assertEqual(len(fake_graph.uploads), 2)
            self.assertEqual(fake_graph.uploads[0]["expected_etag"], "tgt-etag-1")
            self.assertEqual(fake_graph.uploads[1]["expected_etag"], "etag-updated-after-412")
            self.assertIn("upload_result", report)

    def test_retryable_error_retries_successfully(self):
        """Lỗi tạm thời (HTTP 429) retry thành công theo Retry-After."""
        fake_graph = FakeGraphClient(first_upload_429=True)
        fake_graph.set_file(self.cfg.source_path, self.source_bytes, etag="src-etag-1", item_id="src-1")
        fake_graph.set_file(self.cfg.target_path, self.target_bytes, etag="tgt-etag-1", item_id="tgt-1")

        with tempfile.TemporaryDirectory() as tmpdir:
            report = run_nvl_sync(self.cfg, out_dir=tmpdir, publish=True, graph=fake_graph, max_publish_attempts=2)
            self.assertEqual(fake_graph.upload_attempts, 2)
            self.assertIn("upload_result", report)

    def test_offline_mode_cannot_combine_with_publish(self):
        """Chế độ offline từ chối kết hợp với cờ --publish."""
        with tempfile.TemporaryDirectory() as tmpdir:
            src_file = Path(tmpdir) / "src.xlsm"
            tgt_file = Path(tmpdir) / "tgt.xlsx"
            src_file.write_bytes(self.source_bytes)
            tgt_file.write_bytes(self.target_bytes)

            with self.assertRaises(ValueError) as ctx:
                run_nvl_sync(self.cfg, source_file=src_file, target_file=tgt_file, publish=True)
            self.assertIn("--publish", str(ctx.exception))

    def test_offline_mode_generates_proposal_and_report(self):
        """Chế độ offline đọc file local và xuất proposal + report chính xác."""
        with tempfile.TemporaryDirectory() as tmpdir:
            src_file = Path(tmpdir) / "src.xlsm"
            tgt_file = Path(tmpdir) / "tgt.xlsx"
            out_dir = Path(tmpdir) / "out"
            src_file.write_bytes(self.source_bytes)
            tgt_file.write_bytes(self.target_bytes)

            report = run_nvl_sync(self.cfg, source_file=src_file, target_file=tgt_file, out_dir=out_dir, publish=False)
            self.assertEqual(report["mode"], "offline")
            self.assertEqual(report["metrics"]["changed_count"], 2)
            self.assertTrue((out_dir / "nvl_stock_proposal.xlsx").exists())
            self.assertTrue((out_dir / "nvl_stock_report.json").exists())


if __name__ == "__main__":
    unittest.main()
