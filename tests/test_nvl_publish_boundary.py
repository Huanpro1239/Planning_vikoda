"""Integration & boundary tests for sync_nvl_stock using mock Graph."""

from io import BytesIO
import json
from pathlib import Path
import tempfile
from typing import Any
import unittest
from unittest.mock import patch

from openpyxl import load_workbook

from sync_nvl_stock import run_nvl_sync
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

    def test_post_upload_verify_value_mismatch_fails_without_rollback_or_reupload(self):
        """[P1] Xác minh sau upload thấy D2 bị đổi (ví dụ 777): dừng ngay, không published, không upload lại/rollback."""
        corrupted_bytes = make_mock_target_bytes([("VT001", 777), ("VT002", 70)])

        class WrongPostUploadClient(FakeGraphClient):
            def upload_file(self, drive_id, item_id, content, expected_etag=None):
                res = super().upload_file(drive_id, item_id, content, expected_etag)
                # Giả lập file trên server bị thay đổi đồng thời ngay sau upload
                self.files[item_id] = corrupted_bytes
                return res

        fake_graph = WrongPostUploadClient()
        fake_graph.set_file(self.cfg.source_path, self.source_bytes, etag="src-1", item_id="src-1")
        fake_graph.set_file(self.cfg.target_path, self.target_bytes, etag="tgt-1", item_id="tgt-1")

        with tempfile.TemporaryDirectory() as tmpdir:
            with self.assertRaises(RuntimeError) as ctx:
                run_nvl_sync(self.cfg, out_dir=tmpdir, publish=True, graph=fake_graph, max_publish_attempts=3)
            self.assertIn("xác minh sau upload thất bại", str(ctx.exception).lower())

            # Không upload lại: số lượt upload duy nhất là 1
            self.assertEqual(len(fake_graph.uploads), 1)

            # Không rollback: file trên server giữ nguyên giá trị người khác đã sửa (777)
            wb = load_workbook(BytesIO(fake_graph.files["tgt-1"]), data_only=True)
            self.assertEqual(wb["Ton_NVL"]["D2"].value, 777)
            wb.close()

            # Proposal không còn tồn tại để tránh hiểu lầm
            self.assertFalse((Path(tmpdir) / "nvl_stock_proposal.xlsx").exists())

            # Báo cáo failed với phase post_upload_verify và xác nhận upload đã từng thành công
            report_file = Path(tmpdir) / "nvl_stock_report.json"
            self.assertTrue(report_file.exists())
            report = json.loads(report_file.read_text(encoding="utf-8"))
            self.assertEqual(report["status"], "failed")
            self.assertEqual(report["phase"], "post_upload_verify")
            self.assertTrue(report.get("upload_acknowledged"))
            self.assertEqual(report.get("verification_status"), "conflict_or_mismatch")

    def test_post_upload_verify_other_cell_modified_concurrently_fails_without_reupload(self):
        """[P1] Xác minh sau upload thấy ô ngoài D bị sửa đồng thời: báo lỗi, giữ nguyên file server, không re-upload."""
        corrupted_bytes = make_mock_target_bytes([("CONCURRENT_USER", 100), ("VT002", 70)])

        class ConcurrentEditClient(FakeGraphClient):
            def upload_file(self, drive_id, item_id, content, expected_etag=None):
                res = super().upload_file(drive_id, item_id, content, expected_etag)
                self.files[item_id] = corrupted_bytes
                return res

        fake_graph = ConcurrentEditClient()
        fake_graph.set_file(self.cfg.source_path, self.source_bytes, etag="src-1", item_id="src-1")
        fake_graph.set_file(self.cfg.target_path, self.target_bytes, etag="tgt-1", item_id="tgt-1")

        with tempfile.TemporaryDirectory() as tmpdir:
            with self.assertRaises(RuntimeError):
                run_nvl_sync(self.cfg, out_dir=tmpdir, publish=True, graph=fake_graph)

            self.assertEqual(len(fake_graph.uploads), 1)
            self.assertFalse((Path(tmpdir) / "nvl_stock_proposal.xlsx").exists())

            report = json.loads((Path(tmpdir) / "nvl_stock_report.json").read_text(encoding="utf-8"))
            self.assertEqual(report["status"], "failed")
            self.assertEqual(report["phase"], "post_upload_verify")
            self.assertTrue(report.get("upload_acknowledged"))

    def test_post_upload_verify_transient_download_error_retries_and_succeeds(self):
        """[P1] Lỗi GET tạm thời khi tải lại để xác minh: retry bước đọc có giới hạn và thành công, không upload lại."""
        class TransientVerifyDownloadClient(FakeGraphClient):
            def __init__(self):
                super().__init__()
                self.verify_download_attempts = 0

            def download_file(self, drive_id, item_id):
                if self.uploads:
                    self.verify_download_attempts += 1
                    if self.verify_download_attempts == 1:
                        err = GraphRequestError("Service Unavailable", status_code=503)
                        err.retry_after_seconds = 0.01
                        raise err
                return super().download_file(drive_id, item_id)

        fake_graph = TransientVerifyDownloadClient()
        fake_graph.set_file(self.cfg.source_path, self.source_bytes, etag="src-1", item_id="src-1")
        fake_graph.set_file(self.cfg.target_path, self.target_bytes, etag="tgt-1", item_id="tgt-1")

        with tempfile.TemporaryDirectory() as tmpdir:
            report = run_nvl_sync(self.cfg, out_dir=tmpdir, publish=True, graph=fake_graph)
            self.assertEqual(report["status"], "published")
            self.assertTrue(report.get("post_upload_verified"))
            # Duy nhất 1 lần upload, bước đọc verify được thử 2 lần
            self.assertEqual(len(fake_graph.uploads), 1)
            self.assertEqual(fake_graph.verify_download_attempts, 2)

    def test_post_upload_verify_transient_download_error_exhausted_fails_with_report(self):
        """[P1] Lỗi GET tạm thời khi xác minh sau upload hết lượt thử: báo trạng thái unverified, không published."""
        class PersistentVerifyFailClient(FakeGraphClient):
            def __init__(self):
                super().__init__()
                self.verify_download_attempts = 0

            def download_file(self, drive_id, item_id):
                if self.uploads:
                    self.verify_download_attempts += 1
                    err = GraphRequestError("Service Unavailable", status_code=503)
                    err.retry_after_seconds = 0.01
                    raise err
                return super().download_file(drive_id, item_id)

        fake_graph = PersistentVerifyFailClient()
        fake_graph.set_file(self.cfg.source_path, self.source_bytes, etag="src-1", item_id="src-1")
        fake_graph.set_file(self.cfg.target_path, self.target_bytes, etag="tgt-1", item_id="tgt-1")

        with tempfile.TemporaryDirectory() as tmpdir:
            with self.assertRaises(RuntimeError) as ctx:
                run_nvl_sync(self.cfg, out_dir=tmpdir, publish=True, graph=fake_graph)
            self.assertIn("không thể tải lại file để xác minh", str(ctx.exception).lower())

            # Không upload lại
            self.assertEqual(len(fake_graph.uploads), 1)
            self.assertEqual(fake_graph.verify_download_attempts, 3)

            report = json.loads((Path(tmpdir) / "nvl_stock_report.json").read_text(encoding="utf-8"))
            self.assertEqual(report["status"], "failed")
            self.assertEqual(report["phase"], "post_upload_verify")
            self.assertEqual(report.get("verification_status"), "unverified")
            self.assertTrue(report.get("upload_acknowledged"))
            self.assertFalse((Path(tmpdir) / "nvl_stock_proposal.xlsx").exists())

    def test_source_freshness_continuous_churn_after_download_produces_error_report(self):
        """[P2] Nguồn thay đổi liên tục ngay sau download: hết lượt thử thì báo lỗi, ghi report và không upload."""
        class ChurnAfterDownloadClient(FakeGraphClient):
            def __init__(self, src_path):
                super().__init__()
                self.src_path = src_path
                self.dl_count = 0

            def download_file(self, drive_id, item_id):
                data = super().download_file(drive_id, item_id)
                if item_id == "src-1":
                    self.dl_count += 1
                    self.items[self.src_path]["eTag"] = f"src-v{self.dl_count}-changed"
                return data

        fake_graph = ChurnAfterDownloadClient(self.cfg.source_path)
        fake_graph.set_file(self.cfg.source_path, self.source_bytes, etag="src-v0", item_id="src-1")
        fake_graph.set_file(self.cfg.target_path, self.target_bytes, etag="tgt-v0", item_id="tgt-1")

        with tempfile.TemporaryDirectory() as tmpdir, patch("sync_nvl_stock.time.sleep"):
            with self.assertRaises(RuntimeError) as ctx:
                run_nvl_sync(self.cfg, out_dir=tmpdir, publish=True, graph=fake_graph, max_publish_attempts=2)
            self.assertIn("nguồn đã thay đổi", str(ctx.exception).lower())

            self.assertEqual(len(fake_graph.uploads), 0)
            self.assertFalse((Path(tmpdir) / "nvl_stock_proposal.xlsx").exists())

            report = json.loads((Path(tmpdir) / "nvl_stock_report.json").read_text(encoding="utf-8"))
            self.assertEqual(report["status"], "failed")
            self.assertEqual(report["phase"], "verify_source_freshness")
            self.assertEqual(report["attempt"], 2)

    def test_source_freshness_churn_before_upload_produces_error_report(self):
        """[P2] Nguồn thay đổi liên tục trước upload: hết lượt thử thì báo lỗi, ghi report và không upload."""
        class ChurnBeforeUploadClient(FakeGraphClient):
            def __init__(self, src_path):
                super().__init__()
                self.src_path = src_path
                self.calls = 0

            def get_item_by_path(self, drive_id, path):
                item = super().get_item_by_path(drive_id, path)
                if path == self.src_path:
                    self.calls += 1
                    # Cứ mỗi lần pre_upload_source_check (lần gọi thứ 3, 6) thì eTag đổi
                    if self.calls % 3 == 0:
                        item["eTag"] = f"src-churned-pre-{self.calls}"
                return item

        fake_graph = ChurnBeforeUploadClient(self.cfg.source_path)
        fake_graph.set_file(self.cfg.source_path, self.source_bytes, etag="src-base", item_id="src-1")
        fake_graph.set_file(self.cfg.target_path, self.target_bytes, etag="tgt-base", item_id="tgt-1")

        with tempfile.TemporaryDirectory() as tmpdir, patch("sync_nvl_stock.time.sleep"):
            with self.assertRaises(RuntimeError) as ctx:
                run_nvl_sync(self.cfg, out_dir=tmpdir, publish=True, graph=fake_graph, max_publish_attempts=2)
            self.assertIn("nguồn đã bị thay đổi trước khi upload", str(ctx.exception).lower())

            self.assertEqual(len(fake_graph.uploads), 0)
            self.assertFalse((Path(tmpdir) / "nvl_stock_proposal.xlsx").exists())

            report = json.loads((Path(tmpdir) / "nvl_stock_report.json").read_text(encoding="utf-8"))
            self.assertEqual(report["status"], "failed")
            self.assertEqual(report["phase"], "pre_upload_source_check")
            self.assertEqual(report["attempt"], 2)

    def test_config_missing_target_path_produces_error_report(self):
        """[P2] Thiếu target_path sinh lỗi ValueError và xuất error report failed (P2)."""
        self.cfg.target_path = ""
        fake_graph = FakeGraphClient()

        with tempfile.TemporaryDirectory() as tmpdir:
            with self.assertRaises(ValueError):
                run_nvl_sync(self.cfg, out_dir=tmpdir, graph=fake_graph)

            report_file = Path(tmpdir) / "nvl_stock_report.json"
            self.assertTrue(report_file.exists())
            report = json.loads(report_file.read_text(encoding="utf-8"))
            self.assertEqual(report["status"], "failed")
            self.assertEqual(report["phase"], "validate_config")

    def test_config_missing_source_path_produces_error_report(self):
        """[P2] Thiếu source_path sinh lỗi ValueError và xuất error report failed (P2)."""
        self.cfg.source_path = ""
        fake_graph = FakeGraphClient()

        with tempfile.TemporaryDirectory() as tmpdir:
            with self.assertRaises(ValueError):
                run_nvl_sync(self.cfg, out_dir=tmpdir, graph=fake_graph)

            report_file = Path(tmpdir) / "nvl_stock_report.json"
            self.assertTrue(report_file.exists())
            report = json.loads(report_file.read_text(encoding="utf-8"))
            self.assertEqual(report["status"], "failed")
            self.assertEqual(report["phase"], "validate_config")

    # =========================================================================
    # BỘ REGRESSION TESTS REVIEW VÒNG 3: TIMEOUT UPLOAD VÀ VÒNG ĐỜI ARTIFACT CLI
    # =========================================================================

    def test_timeout_server_committed_concurrent_edit_in_d_stops_without_reupload(self):
        """[Vòng 3 - P1] Timeout upload nhưng server đã ghi và D bị sửa đồng thời:
        Dừng ngay, không retry PUT với ETag mới, không đè 777 về 100, chỉ upload 1 lần."""
        class TimeoutConcurrentEditD(FakeGraphClient):
            def __init__(self, target_path: str):
                super().__init__()
                self.target_path = target_path

            def upload_file(self, drive_id, item_id, content, expected_etag=None):
                super().upload_file(drive_id, item_id, content, expected_etag)
                if self.upload_attempts == 1:
                    # Giả lập server đã nhận nhưng socket timeout; đồng thời D2 bị sửa thành 777
                    self.files[item_id] = make_mock_target_bytes([("VT001", 777), ("VT002", 70)])
                    self.items[self.target_path]["eTag"] = "etag-user-edit-777"
                    raise TimeoutError("The write operation timed out waiting for server ack")
                return {"id": item_id, "name": "uploaded.xlsx", "eTag": "etag-new"}

        fake_graph = TimeoutConcurrentEditD(self.cfg.target_path)
        fake_graph.set_file(self.cfg.source_path, self.source_bytes, etag="src-1", item_id="src-1")
        fake_graph.set_file(self.cfg.target_path, self.target_bytes, etag="tgt-1", item_id="tgt-1")

        with tempfile.TemporaryDirectory() as tmpdir, patch("sync_nvl_stock.time.sleep"):
            with self.assertRaises(RuntimeError) as ctx:
                run_nvl_sync(self.cfg, out_dir=tmpdir, publish=True, graph=fake_graph, max_publish_attempts=3)
            self.assertIn("xác minh sau đó thất bại", str(ctx.exception).lower())

            # Tuyệt đối chỉ upload 1 lần, không lặp lại với ETag mới
            self.assertEqual(fake_graph.upload_attempts, 1)

            # Dữ liệu 777 của người dùng trên server được bảo toàn nguyên vẹn
            wb = load_workbook(BytesIO(fake_graph.files["tgt-1"]), data_only=True)
            self.assertEqual(wb["Ton_NVL"]["D2"].value, 777)
            wb.close()

            report = json.loads((Path(tmpdir) / "nvl_stock_report.json").read_text(encoding="utf-8"))
            self.assertEqual(report["status"], "failed")
            self.assertEqual(report["phase"], "post_upload_verify")
            self.assertFalse(report.get("upload_acknowledged"))
            self.assertTrue(report.get("upload_may_have_committed"))
            self.assertEqual(report.get("verification_status"), "conflict_or_mismatch")

    def test_timeout_server_committed_concurrent_edit_outside_d_stops_without_reupload(self):
        """[Vòng 3 - P1] Timeout upload nhưng server đã ghi và ô ngoài cột D bị sửa đồng thời:
        Dừng ngay, chỉ upload 1 lần, bảo toàn dữ liệu server."""
        class TimeoutConcurrentEditOutsideD(FakeGraphClient):
            def __init__(self, target_path: str):
                super().__init__()
                self.target_path = target_path

            def upload_file(self, drive_id, item_id, content, expected_etag=None):
                super().upload_file(drive_id, item_id, content, expected_etag)
                if self.upload_attempts == 1:
                    # Giả lập sửa tên vật tư ở cột B hoặc thêm dòng ngoài D
                    corrupted = make_mock_target_bytes([("VT001", 50), ("VT002", 70)])
                    wb_mod = load_workbook(BytesIO(corrupted))
                    wb_mod["Ton_NVL"]["B2"].value = "Tên bị sửa đồng thời"
                    buf = BytesIO()
                    wb_mod.save(buf)
                    wb_mod.close()
                    self.files[item_id] = buf.getvalue()
                    self.items[self.target_path]["eTag"] = "etag-outside-d-changed"
                    raise TimeoutError("Socket timed out")
                return {"id": item_id, "name": "uploaded.xlsx", "eTag": "etag-new"}

        fake_graph = TimeoutConcurrentEditOutsideD(self.cfg.target_path)
        fake_graph.set_file(self.cfg.source_path, self.source_bytes, etag="src-1", item_id="src-1")
        fake_graph.set_file(self.cfg.target_path, self.target_bytes, etag="tgt-1", item_id="tgt-1")

        with tempfile.TemporaryDirectory() as tmpdir, patch("sync_nvl_stock.time.sleep"):
            with self.assertRaises(RuntimeError) as ctx:
                run_nvl_sync(self.cfg, out_dir=tmpdir, publish=True, graph=fake_graph, max_publish_attempts=3)
            self.assertIn("xác minh sau đó thất bại", str(ctx.exception).lower())

            self.assertEqual(fake_graph.upload_attempts, 1)
            report = json.loads((Path(tmpdir) / "nvl_stock_report.json").read_text(encoding="utf-8"))
            self.assertEqual(report["status"], "failed")
            self.assertEqual(report["phase"], "post_upload_verify")
            self.assertFalse(report.get("upload_acknowledged"))
            self.assertTrue(report.get("upload_may_have_committed"))
            self.assertEqual(report.get("verification_status"), "conflict_or_mismatch")

    def test_timeout_get_transient_error_recovers_and_publishes(self):
        """[Vòng 3 - P1] Timeout upload nhưng server đã ghi; GET verify gặp lỗi mạng tạm thời (503/timeout):
        Retry GET thành công, xác minh đạt và hoàn tất publish với 1 lượt upload duy nhất."""
        class TimeoutRecoveringClient(FakeGraphClient):
            def __init__(self):
                super().__init__(timeout_then_server_committed=True)
                self.post_download_calls = 0

            def download_file(self, drive_id, item_id):
                # Khi đang download để verify sau upload (item_id là tgt-1 sau khi upload_attempts >= 1)
                if self.upload_attempts >= 1 and item_id == "tgt-1":
                    self.post_download_calls += 1
                    if self.post_download_calls == 1:
                        err = GraphRequestError("Service Unavailable", status_code=503, error_code="serviceUnavailable")
                        err.retry_after_seconds = 0.01
                        raise err
                return super().download_file(drive_id, item_id)

        fake_graph = TimeoutRecoveringClient()
        fake_graph.set_file(self.cfg.source_path, self.source_bytes, etag="src-1", item_id="src-1")
        fake_graph.set_file(self.cfg.target_path, self.target_bytes, etag="tgt-1", item_id="tgt-1")

        with tempfile.TemporaryDirectory() as tmpdir, patch("sync_nvl_stock.time.sleep"):
            report = run_nvl_sync(self.cfg, out_dir=tmpdir, publish=True, graph=fake_graph, max_publish_attempts=3)
            self.assertEqual(report["status"], "published")
            self.assertEqual(fake_graph.upload_attempts, 1)
            self.assertTrue(report.get("post_upload_verified"))

    def test_timeout_get_exhausted_stops_as_unverified_without_reupload(self):
        """[Vòng 3 - P1] Timeout upload nhưng GET verify hết lượt thử (lỗi mạng liên tục):
        Dừng ngay với status unverified, ghi nhận upload có thể đã thành công, không upload lại."""
        class TimeoutGetExhaustedClient(FakeGraphClient):
            def __init__(self):
                super().__init__(timeout_then_server_committed=True)

            def download_file(self, drive_id, item_id):
                if self.upload_attempts >= 1:
                    raise TimeoutError("Network timeout during verification download")
                return super().download_file(drive_id, item_id)

        fake_graph = TimeoutGetExhaustedClient()
        fake_graph.set_file(self.cfg.source_path, self.source_bytes, etag="src-1", item_id="src-1")
        fake_graph.set_file(self.cfg.target_path, self.target_bytes, etag="tgt-1", item_id="tgt-1")

        with tempfile.TemporaryDirectory() as tmpdir, patch("sync_nvl_stock.time.sleep"):
            with self.assertRaises(RuntimeError) as ctx:
                run_nvl_sync(self.cfg, out_dir=tmpdir, publish=True, graph=fake_graph, max_publish_attempts=3)
            self.assertIn("không thể tải lại file để xác minh", str(ctx.exception).lower())

            self.assertEqual(fake_graph.upload_attempts, 1)
            report = json.loads((Path(tmpdir) / "nvl_stock_report.json").read_text(encoding="utf-8"))
            self.assertEqual(report["status"], "failed")
            self.assertEqual(report["phase"], "post_upload_verify")
            self.assertFalse(report.get("upload_acknowledged"))
            self.assertTrue(report.get("upload_may_have_committed"))
            self.assertEqual(report.get("verification_status"), "unverified")

    def test_backup_write_failure_stops_with_zero_upload_attempts(self):
        """Khi lưu hoặc xác minh backup thực tế thất bại trước upload:
        Dừng ngay lập tức với report failed phase=pre_upload_backup và KHÔNG gọi upload (upload_attempts=0)."""
        fake_graph = FakeGraphClient()
        fake_graph.set_file(self.cfg.source_path, self.source_bytes, etag="src-1", item_id="src-1")
        fake_graph.set_file(self.cfg.target_path, self.target_bytes, etag="tgt-1", item_id="tgt-1")

        with tempfile.TemporaryDirectory() as tmpdir:
            out_dir = Path(tmpdir)
            with patch.object(Path, "write_bytes", side_effect=OSError("Disk full: không thể ghi backup raw")):
                with self.assertRaises(RuntimeError) as ctx:
                    run_nvl_sync(self.cfg, out_dir=tmpdir, publish=True, graph=fake_graph, max_publish_attempts=3)
                self.assertIn("backup thực tế đích thất bại", str(ctx.exception).lower())

            # BẮT BUỘC: upload_attempts phải bằng 0!
            self.assertEqual(fake_graph.upload_attempts, 0)

            # Report phải ghi nhận trạng thái failed tại phase pre_upload_backup
            rep_file = out_dir / "nvl_stock_report.json"
            self.assertTrue(rep_file.exists())
            rep = json.loads(rep_file.read_text(encoding="utf-8"))
            self.assertEqual(rep["status"], "failed")
            self.assertEqual(rep["phase"], "pre_upload_backup")
            self.assertIn("Disk full", rep["error_message"])

    def test_cli_missing_config_cleans_stale_proposal_and_writes_failed_report(self):
        """[Vòng 3 - P2] Chạy CLI với file cấu hình không tồn tại:
        Xóa proposal cũ của lần chạy trước, ghi đè report cũ bằng status=failed, phase=init_cli."""
        with tempfile.TemporaryDirectory() as tmpdir:
            stale_report = Path(tmpdir) / "nvl_stock_report.json"
            stale_report.write_text('{"status":"published","message":"old success"}', encoding="utf-8")
            stale_proposal = Path(tmpdir) / "nvl_stock_proposal.xlsx"
            stale_proposal.write_bytes(b"old-proposal-content")

            import subprocess, sys
            proc = subprocess.run(
                [sys.executable, "-X", "utf8", "sync_nvl_stock.py", "--config", str(Path(tmpdir) / "non_existent.json"), "--out", tmpdir],
                capture_output=True,
                text=True,
                encoding="utf-8",
            )
            self.assertNotEqual(proc.returncode, 0)

            # Proposal cũ phải bị xóa sạch
            self.assertFalse(stale_proposal.exists())

            # Report phải được ghi đè bằng lỗi init_cli hiện tại
            report = json.loads(stale_report.read_text(encoding="utf-8"))
            self.assertEqual(report["status"], "failed")
            self.assertEqual(report["phase"], "init_cli")
            self.assertEqual(report["error_type"], "FileNotFoundError")

    def test_cli_corrupt_json_config_cleans_stale_proposal_and_writes_failed_report(self):
        """[Vòng 3 - P2] Chạy CLI với file JSON hỏng:
        Xóa proposal cũ, ghi đè report cũ bằng status=failed, phase=init_cli."""
        with tempfile.TemporaryDirectory() as tmpdir:
            corrupt_cfg = Path(tmpdir) / "corrupt.json"
            corrupt_cfg.write_text('{"invalid_json": [missing_bracket', encoding="utf-8")

            stale_report = Path(tmpdir) / "nvl_stock_report.json"
            stale_report.write_text('{"status":"published","message":"old success"}', encoding="utf-8")
            stale_proposal = Path(tmpdir) / "nvl_stock_proposal.xlsx"
            stale_proposal.write_bytes(b"old-proposal-content")

            import subprocess, sys
            proc = subprocess.run(
                [sys.executable, "-X", "utf8", "sync_nvl_stock.py", "--config", str(corrupt_cfg), "--out", tmpdir],
                capture_output=True,
                text=True,
                encoding="utf-8",
            )
            self.assertNotEqual(proc.returncode, 0)

            self.assertFalse(stale_proposal.exists())
            report = json.loads(stale_report.read_text(encoding="utf-8"))
            self.assertEqual(report["status"], "failed")
            self.assertEqual(report["phase"], "init_cli")

    def test_cli_post_upload_verify_failure_preserves_detailed_report(self):
        """[Vòng 3 - P2] Khi run_nvl_sync thất bại ở post_upload_verify, CLI không ghi đè làm mất báo cáo lỗi chi tiết."""
        corrupted_bytes = make_mock_target_bytes([("VT001", 777), ("VT002", 70)])

        class MismatchClient(FakeGraphClient):
            def upload_file(self, drive_id, item_id, content, expected_etag=None):
                res = super().upload_file(drive_id, item_id, content, expected_etag)
                self.files[item_id] = corrupted_bytes
                return res

        fake_graph = MismatchClient()
        fake_graph.set_file(self.cfg.source_path, self.source_bytes, etag="src-1", item_id="src-1")
        fake_graph.set_file(self.cfg.target_path, self.target_bytes, etag="tgt-1", item_id="tgt-1")

        with tempfile.TemporaryDirectory() as tmpdir:
            cfg_file = Path(tmpdir) / "config.json"
            cfg_file.write_text(
                json.dumps({
                    "source": {"name": "src", "sharepoint_path": self.cfg.source_path, "sourcedoc": "SRC"},
                    "target": {"name": "tgt", "sharepoint_path": self.cfg.target_path, "sourcedoc": "TGT"},
                }),
                encoding="utf-8",
            )

            # Gọi run_nvl_sync với mismatch để xác nhận report lưu đúng phase post_upload_verify
            with self.assertRaises(RuntimeError):
                run_nvl_sync(self.cfg, out_dir=tmpdir, publish=True, graph=fake_graph)

            report = json.loads((Path(tmpdir) / "nvl_stock_report.json").read_text(encoding="utf-8"))
            self.assertEqual(report["status"], "failed")
            self.assertEqual(report["phase"], "post_upload_verify")
            self.assertEqual(report.get("verification_status"), "conflict_or_mismatch")


if __name__ == "__main__":
    unittest.main()
