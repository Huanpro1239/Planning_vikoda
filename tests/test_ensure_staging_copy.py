import json
import shutil
import tempfile
import unittest
from pathlib import Path
from unittest.mock import MagicMock, patch

from scripts.ensure_staging_copy import ensure_staging_copy, parse_sourcedoc_from_etag
from sync_stock import GraphClient, GraphRequestError


class EnsureStagingCopyTests(unittest.TestCase):
    def setUp(self):
        self.tmp_dir = tempfile.TemporaryDirectory()
        self.workdir = Path(self.tmp_dir.name)

        self.staging_cfg_path = self.workdir / "staging_config.json"
        staging_data = {
            "source": {
                "name": "XNT_ketoan_Vikoda.xlsm",
                "sharepoint_path": "Tinh san xuat Mua hang 2027/Ton He thong/Ton Ke Toan/XNT_ketoan_Vikoda.xlsm",
                "sourcedoc": "C0372C81-A402-4A11-B9E0-847768CC9CFB",
            },
            "target": {
                "name": "Test_Ke_hoach_mua_hang_copy.xlsx",
                "sharepoint_path": "Tinh san xuat Mua hang 2027/Test_Ke_hoach_mua_hang_copy.xlsx",
                "sourcedoc": "",
            },
        }
        self.staging_cfg_path.write_text(json.dumps(staging_data, indent=2), encoding="utf-8")

        self.official_cfg_path = self.workdir / "official_config.json"
        official_data = {
            "source": {
                "name": "XNT_ketoan_Vikoda.xlsm",
                "sharepoint_path": "Tinh san xuat Mua hang 2027/Ton He thong/Ton Ke Toan/XNT_ketoan_Vikoda.xlsm",
                "sourcedoc": "C0372C81-A402-4A11-B9E0-847768CC9CFB",
            },
            "target": {
                "name": "Kế hoạch mua hàng.xlsx",
                "sharepoint_path": "Tinh san xuat Mua hang 2027/Kế hoạch mua hàng.xlsx",
                "sourcedoc": "89D1BA7B-006E-4527-B879-ABF120309214",
            },
        }
        self.official_cfg_path.write_text(json.dumps(official_data, indent=2), encoding="utf-8")

        self.out_dir = self.workdir / "out"

    def tearDown(self):
        self.tmp_dir.cleanup()

    def test_parse_sourcedoc_from_etag(self):
        etag1 = '"{89D1BA7B-006E-4527-B879-ABF120309214},20"'
        self.assertEqual(parse_sourcedoc_from_etag(etag1), "89D1BA7B-006E-4527-B879-ABF120309214")

        etag2 = '"{abc-123},1"'
        self.assertEqual(parse_sourcedoc_from_etag(etag2), "ABC-123")

        self.assertEqual(parse_sourcedoc_from_etag(""), "")
        self.assertEqual(parse_sourcedoc_from_etag("invalid_etag"), "")

    def test_ensure_staging_copy_when_already_exists(self):
        mock_graph = MagicMock(spec=GraphClient)
        mock_graph.get_site_id.return_value = "site_123"
        mock_graph.get_default_drive_id.return_value = "drive_123"

        existing_item = {
            "id": "item_copy_existing",
            "name": "Test_Ke_hoach_mua_hang_copy.xlsx",
            "eTag": '"{11111111-2222-3333-4444-555555555555},5"',
            "size": 54321,
            "webUrl": "https://sharepoint.com/test_copy.xlsx",
        }
        mock_graph.get_item_by_path.return_value = existing_item

        res = ensure_staging_copy(
            config_path=str(self.staging_cfg_path),
            official_config_path=str(self.official_cfg_path),
            out_dir_path=str(self.out_dir),
            graph=mock_graph,
        )

        self.assertEqual(res["status"], "existed")
        self.assertEqual(res["sourcedoc"], "11111111-2222-3333-4444-555555555555")
        self.assertEqual(res["item_id"], "item_copy_existing")

        # Đảm bảo không tải file gốc và không gọi tạo file
        mock_graph.download_file.assert_not_called()
        mock_graph.create_file_by_path.assert_not_called()

        # Kiểm tra file output
        info_file = self.out_dir / "staging_copy_info.json"
        self.assertTrue(info_file.exists())
        info_data = json.loads(info_file.read_text(encoding="utf-8"))
        self.assertEqual(info_data["status"], "existed")
        self.assertEqual(info_data["sourcedoc"], "11111111-2222-3333-4444-555555555555")

        report_file = self.out_dir / "nvl_stock_report.json"
        self.assertTrue(report_file.exists())
        report_data = json.loads(report_file.read_text(encoding="utf-8"))
        self.assertEqual(report_data["mode"], "ensure_staging_copy")
        self.assertEqual(report_data["status"], "success")

    def test_ensure_staging_copy_when_not_exists_creates_new(self):
        mock_graph = MagicMock(spec=GraphClient)
        mock_graph.get_site_id.return_value = "site_123"
        mock_graph.get_default_drive_id.return_value = "drive_123"

        official_item = {
            "id": "item_official_original",
            "name": "Kế hoạch mua hàng.xlsx",
            "eTag": '"{89D1BA7B-006E-4527-B879-ABF120309214},20"',
            "size": 50000,
        }
        new_copy_item = {
            "id": "item_copy_newly_created",
            "name": "Test_Ke_hoach_mua_hang_copy.xlsx",
            "eTag": '"{99999999-8888-7777-6666-555555555555},1"',
            "size": 50000,
            "webUrl": "https://sharepoint.com/test_copy.xlsx",
        }

        # Lần 1: check copy path -> 404
        # Lần 2: check official path -> official_item
        # Lần 3: re-fetch copy path -> new_copy_item
        def mock_get_item_by_path(drive_id, path):
            if "Test_Ke_hoach_mua_hang_copy.xlsx" in path:
                if not hasattr(mock_get_item_by_path, "created"):
                    mock_get_item_by_path.created = True
                    raise GraphRequestError("Not found", status_code=404)
                return new_copy_item
            elif "Kế hoạch mua hàng.xlsx" in path:
                return official_item
            raise GraphRequestError("Unknown path", status_code=404)

        mock_graph.get_item_by_path.side_effect = mock_get_item_by_path
        mock_graph.download_file.return_value = b"official_bytes_content"
        mock_graph.create_file_by_path.return_value = new_copy_item

        res = ensure_staging_copy(
            config_path=str(self.staging_cfg_path),
            official_config_path=str(self.official_cfg_path),
            out_dir_path=str(self.out_dir),
            graph=mock_graph,
        )

        self.assertEqual(res["status"], "created")
        self.assertEqual(res["sourcedoc"], "99999999-8888-7777-6666-555555555555")
        self.assertEqual(res["item_id"], "item_copy_newly_created")

        mock_graph.download_file.assert_called_once_with("drive_123", "item_official_original")
        mock_graph.create_file_by_path.assert_called_once_with(
            "drive_123",
            "Tinh san xuat Mua hang 2027/Test_Ke_hoach_mua_hang_copy.xlsx",
            b"official_bytes_content",
            conflict_behavior="fail",
        )

        # Kiểm tra file output
        info_file = self.out_dir / "staging_copy_info.json"
        self.assertTrue(info_file.exists())
        info_data = json.loads(info_file.read_text(encoding="utf-8"))
        self.assertEqual(info_data["status"], "created")
        self.assertEqual(info_data["sourcedoc"], "99999999-8888-7777-6666-555555555555")

    def test_graph_client_create_file_by_path(self):
        token = "fake_token"
        client = GraphClient(token)
        mock_resp = MagicMock()
        mock_resp.ok = True
        mock_resp.status_code = 201
        mock_resp.json.return_value = {
            "id": "new_item_id",
            "name": "Test_Ke_hoach_mua_hang_copy.xlsx",
            "eTag": '"{FFFFFFFF-EEEE-DDDD-CCCC-BBBBBBBBBBBB},1"',
        }

        with patch.object(client.session, "put", return_value=mock_resp) as mock_put:
            res = client.create_file_by_path(
                drive_id="drive_1",
                file_path="Tinh san xuat/copy.xlsx",
                content=b"test_data",
                conflict_behavior="fail",
            )
            self.assertEqual(res["id"], "new_item_id")
            mock_put.assert_called_once()
            call_kwargs = mock_put.call_args[1]
            self.assertEqual(call_kwargs["params"], {"@microsoft.graph.conflictBehavior": "fail"})
            self.assertEqual(call_kwargs["data"], b"test_data")


if __name__ == "__main__":
    unittest.main()
