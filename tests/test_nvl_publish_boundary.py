"""Integration & boundary tests for sync_nvl_stock using mock Graph."""

from io import BytesIO
import json
from pathlib import Path
import tempfile
from typing import Any
import unittest

from openpyxl import load_workbook

from sync_nvl_stock import NVLConfig, run_nvl_sync
from sync_stock import GraphRequestError
from tests.test_sync_nvl_stock import make_mock_config, make_mock_source_bytes, make_mock_target_bytes


class FakeGraphClient:
    def __init__(
        self,
        *,
        first_upload_412=False,
        first_upload_429=False,
        always_upload_412=False,
        upload_error=None,
        first_download_429=False,
        timeout_then_server_committed=False,
    ):
        self.site_id = "fake-site-id"
        self.drive_id = "fake-drive-id"
        self.first_upload_412 = first_upload_412
        self.always_upload_412 = always_upload_412
        self.first_upload_429 = first_upload_429
        self.first_download_429 = first_download_429
        self.timeout_then_server_committed = timeout_then_server_committed
        self.upload_error = upload_error

        self.upload_attempts = 0
        self.download_attempts = 0
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
        self.download_attempts += 1
        if self.first_download_429 and self.download_attempts == 1:
            err = GraphRequestError("Too Many Requests", status_code=429, error_code="tooManyRequests")
            err.retry_after_seconds = 0.05
            raise err

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

        if self.upload_error is not None:
            raise self.upload_error

        if self.always_upload_412:
            raise GraphRequestError("preconditionFailed", status_code=412, error_code="preconditionFailed")

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

        if self.timeout_then_server_committed and self.upload_attempts == 1:
            # Giả lập server đã ghi nhận file nhưng socket client bị timeout trước khi nhận response HTTP 200
            self.files[item_id] = content
            for path, item in self.items.items():
                if item.get("id") == item_id:
                    item["eTag"] = "etag-committed-on-timeout"
            raise TimeoutError("The write operation timed out waiting for server ack")

        self.files[item_id] = content
        for path, item in self.items.items():
            if item.get("id") == item_id:
                item["eTag"] = "etag-new"

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
        """Chế độ publish=True khi có thay đổi gọi upload_file với đúng expected_etag và trả status published."""
        fake_graph = FakeGraphClient()
        fake_graph.set_file(self.cfg.source_path, self.source_bytes, etag="src-etag-1", item_id="src-1")
        fake_graph.set_file(self.cfg.target_path, self.target_bytes, etag="tgt-etag-1", item_id="tgt-1")

        with tempfile.TemporaryDirectory() as tmpdir:
            report = run_nvl_sync(self.cfg, out_dir=tmpdir, publish=True, graph=fake_graph)

            self.assertEqual(report["mode"], "publish")
            self.assertEqual(report["status"], "published")
            self.assertEqual(report["metrics"]["changed_count"], 2)
            self.assertEqual(len(fake_graph.uploads), 1)
            self.assertEqual(fake_graph.uploads[0]["expected_etag"], "tgt-etag-1")
            self.assertIn("upload_result", report)

    def test_publish_without_changes_skips_upload(self):
        """Chế độ publish=True khi dữ liệu đã khớp hoàn toàn (0 thay đổi) sẽ bỏ qua upload."""
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

    def test_source_changed_after_download_triggers_retry_and_uploads_latest(self):
        """Phát hiện nguồn thay đổi trong lúc download snapshot hoặc trước upload; tự làm mới và upload số mới (P1)."""
        class DynamicSourceGraph(FakeGraphClient):
            def download_file(self, drive_id, item_id):
                data = super().download_file(drive_id, item_id)
                if item_id == "src-1":
                    # Cập nhật nguồn trên SharePoint sang giá trị 999 và ETag mới
                    self.set_file(
                        self.items["path/to/XNT_ketoan_Vikoda.xlsm"]["name"],
                        make_mock_source_bytes([("VT001", 999), ("VT002", 200)]),
                        etag="src-v2",
                        item_id="src-1",
                    )
                    self.items["path/to/XNT_ketoan_Vikoda.xlsm"]["eTag"] = "src-v2"
                return data

        g = DynamicSourceGraph()
        g.set_file(self.cfg.source_path, self.source_bytes, etag="src-v1", item_id="src-1")
        g.set_file(self.cfg.target_path, self.target_bytes, etag="tgt-v1", item_id="tgt-1")

        with tempfile.TemporaryDirectory() as tmpdir:
            report = run_nvl_sync(self.cfg, out_dir=tmpdir, publish=True, graph=g, max_publish_attempts=3)
            self.assertEqual(report["status"], "published")
            self.assertEqual(report["source"]["revision"], "src-v2")

            # File upload phải nhận đúng giá trị mới VT001 = 999
            wb = load_workbook(BytesIO(g.uploads[-1]["content"]), data_only=True)
            self.assertEqual(wb["Ton_NVL"]["D2"].value, 999.0)
            wb.close()

    def test_upload_error_403_produces_failed_report(self):
        """Lỗi upload (ví dụ HTTP 403 Forbidden) sinh report có status='failed' và phase='upload_target' (P2)."""
        fake_graph = FakeGraphClient(upload_error=GraphRequestError("Forbidden", status_code=403))
        fake_graph.set_file(self.cfg.source_path, self.source_bytes, etag="src-1", item_id="src-1")
        fake_graph.set_file(self.cfg.target_path, self.target_bytes, etag="tgt-1", item_id="tgt-1")

        with tempfile.TemporaryDirectory() as tmpdir:
            with self.assertRaises(GraphRequestError):
                run_nvl_sync(self.cfg, out_dir=tmpdir, publish=True, graph=fake_graph)

            report_file = Path(tmpdir) / "nvl_stock_report.json"
            self.assertTrue(report_file.exists())
            report = json.loads(report_file.read_text(encoding="utf-8"))
            self.assertEqual(report["status"], "failed")
            self.assertEqual(report["mode"], "publish")
            self.assertEqual(report["phase"], "upload_target")
            self.assertEqual(report["error_type"], "GraphRequestError")
            self.assertIn("Forbidden", report["message"])

    def test_precondition_failed_412_exhausted_produces_failed_report(self):
        """Xung đột 412 liên tục vượt quá max_publish_attempts sinh report status='failed' (P2)."""
        fake_graph = FakeGraphClient(always_upload_412=True)
        fake_graph.set_file(self.cfg.source_path, self.source_bytes, etag="src-1", item_id="src-1")
        fake_graph.set_file(self.cfg.target_path, self.target_bytes, etag="tgt-1", item_id="tgt-1")

        with tempfile.TemporaryDirectory() as tmpdir:
            with self.assertRaises(GraphRequestError):
                run_nvl_sync(self.cfg, out_dir=tmpdir, publish=True, graph=fake_graph, max_publish_attempts=2)

            report_file = Path(tmpdir) / "nvl_stock_report.json"
            self.assertTrue(report_file.exists())
            report = json.loads(report_file.read_text(encoding="utf-8"))
            self.assertEqual(report["status"], "failed")
            self.assertEqual(report["attempt"], 2)

    def test_early_validation_error_cleans_stale_files_and_writes_failed_report(self):
        """Lỗi validation sớm (mã trùng lặp) xóa artifact cũ và ghi nhận report status='failed' (P2)."""
        bad_source = make_mock_source_bytes([("VT001", 10), ("VT001", 20)])
        fake_graph = FakeGraphClient()
        fake_graph.set_file(self.cfg.source_path, bad_source, etag="src-1", item_id="src-1")
        fake_graph.set_file(self.cfg.target_path, self.target_bytes, etag="tgt-1", item_id="tgt-1")

        with tempfile.TemporaryDirectory() as tmpdir:
            # Tạo file giả lập từ lần chạy trước
            stale_prop = Path(tmpdir) / "nvl_stock_proposal.xlsx"
            stale_rep = Path(tmpdir) / "nvl_stock_report.json"
            stale_prop.write_bytes(b"stale-content")
            stale_rep.write_text(json.dumps({"status": "old"}), encoding="utf-8")

            with self.assertRaises(RuntimeError) as ctx:
                run_nvl_sync(self.cfg, out_dir=tmpdir, publish=True, graph=fake_graph)
            self.assertIn("lặp", str(ctx.exception).lower())

            # Proposal cũ phải bị xóa bỏ
            self.assertFalse(stale_prop.exists())
            # Report mới phải ghi nhận failure
            rep = json.loads(stale_rep.read_text(encoding="utf-8"))
            self.assertEqual(rep["status"], "failed")
            self.assertIn("reconcile", rep["phase"])

    def test_transient_network_timeout_during_upload_recovers_if_server_committed(self):
        """Xử lý timeout upload: nếu kiểm tra lại thấy server đã ghi nhận file, xác nhận thành công (P2)."""
        fake_graph = FakeGraphClient(timeout_then_server_committed=True)
        fake_graph.set_file(self.cfg.source_path, self.source_bytes, etag="src-1", item_id="src-1")
        fake_graph.set_file(self.cfg.target_path, self.target_bytes, etag="tgt-1", item_id="tgt-1")

        with tempfile.TemporaryDirectory() as tmpdir:
            report = run_nvl_sync(self.cfg, out_dir=tmpdir, publish=True, graph=fake_graph, max_publish_attempts=1)
            self.assertEqual(report["status"], "published")
            self.assertEqual(report["upload_result"]["status"], "verified_post_timeout")

    def test_transient_error_on_download_retries_successfully(self):
        """Lỗi tạm thời HTTP 429 tại bước download được retry thành công (P2)."""
        fake_graph = FakeGraphClient(first_download_429=True)
        fake_graph.set_file(self.cfg.source_path, self.source_bytes, etag="src-1", item_id="src-1")
        fake_graph.set_file(self.cfg.target_path, self.target_bytes, etag="tgt-1", item_id="tgt-1")

        with tempfile.TemporaryDirectory() as tmpdir:
            report = run_nvl_sync(self.cfg, out_dir=tmpdir, publish=True, graph=fake_graph, max_publish_attempts=2)
            self.assertEqual(report["status"], "published")
            # 1: source fail (429) -> retry: 2: source ok, 3: target ok, 4: post-upload verify ok
            self.assertEqual(fake_graph.download_attempts, 4)


if __name__ == "__main__":
    unittest.main()
