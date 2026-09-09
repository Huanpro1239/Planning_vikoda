"""Unit tests for scripts/trigger_and_download_dryrun.py."""

import io
import json
import tempfile
import unittest
import zipfile
from pathlib import Path
from unittest.mock import MagicMock, patch

from scripts.trigger_and_download_dryrun import (
    dispatch_workflow,
    download_run_artifacts,
    find_run_for_sha,
    get_existing_run_ids,
    main,
    verify_run_execution,
    wait_for_run_completion,
)


class TriggerAndDownloadDryrunTests(unittest.TestCase):
    def test_dispatch_workflow_success(self):
        with patch('requests.post') as mock_post:
            mock_post.return_value.status_code = 204
            # Should complete without error
            dispatch_workflow('fake-token', 'owner/repo', 'wf.yml', 'feat/branch', {'publish': False})
            mock_post.assert_called_once()

    def test_dispatch_workflow_failure_raises_runtime_error(self):
        with patch('requests.post') as mock_post:
            mock_post.return_value.status_code = 404
            mock_post.return_value.text = 'Not Found'
            with self.assertRaises(RuntimeError) as ctx:
                dispatch_workflow('fake-token', 'owner/repo', 'wf.yml', 'feat/branch')
            self.assertIn('Workflow dispatch failed with HTTP 404', str(ctx.exception))

    def test_get_existing_run_ids_returns_id_set(self):
        with patch('requests.get') as mock_get:
            mock_get.return_value.status_code = 200
            mock_get.return_value.json.return_value = {
                'workflow_runs': [{'id': 101}, {'id': 102}, {'id': 103}]
            }
            res = get_existing_run_ids('fake-token', 'owner/repo', 'feat/branch', 'sync-nvl-stock.yml')
            self.assertEqual(res, {101, 102, 103})

    def test_find_run_for_sha_matches_correct_sha(self):
        fake_runs = [
            {'id': 101, 'head_sha': 'old_sha_123', 'head_branch': 'feat/branch', 'name': 'Sync SharePoint NVL Stock'},
            {'id': 102, 'head_sha': 'expected_sha_456', 'head_branch': 'feat/branch', 'name': 'Sync SharePoint NVL Stock', 'status': 'queued'},
        ]
        with patch('requests.get') as mock_get:
            mock_get.return_value.status_code = 200
            mock_get.return_value.json.return_value = {'workflow_runs': fake_runs}

            run = find_run_for_sha('fake-token', 'owner/repo', 'feat/branch', 'sync-nvl-stock.yml', 'expected_sha_456', timeout_seconds=5)
            self.assertEqual(run['id'], 102)
            self.assertEqual(run['head_sha'], 'expected_sha_456')

    def test_find_run_for_sha_ignores_excluded_run_ids(self):
        fake_runs = [
            {'id': 101, 'head_sha': 'same_sha_123', 'head_branch': 'feat/branch', 'name': 'Sync SharePoint NVL Stock', 'status': 'completed'},
            {'id': 102, 'head_sha': 'same_sha_123', 'head_branch': 'feat/branch', 'name': 'Sync SharePoint NVL Stock', 'status': 'queued'},
        ]
        with patch('requests.get') as mock_get:
            mock_get.return_value.status_code = 200
            mock_get.return_value.json.return_value = {'workflow_runs': fake_runs}

            # 101 exists prior to dispatch, so find_run_for_sha must return 102
            run = find_run_for_sha(
                'fake-token',
                'owner/repo',
                'feat/branch',
                'sync-nvl-stock.yml',
                'same_sha_123',
                exclude_run_ids={101},
                timeout_seconds=5,
            )
            self.assertEqual(run['id'], 102)

    def test_find_run_for_sha_requires_event_match(self):
        fake_runs = [
            {'id': 101, 'head_sha': 'sha_123', 'head_branch': 'feat/branch', 'name': 'Sync SharePoint NVL Stock', 'event': 'push'},
            {'id': 102, 'head_sha': 'sha_123', 'head_branch': 'feat/branch', 'name': 'Sync SharePoint NVL Stock', 'event': 'workflow_dispatch'},
        ]
        with patch('requests.get') as mock_get:
            mock_get.return_value.status_code = 200
            mock_get.return_value.json.return_value = {'workflow_runs': fake_runs}

            run = find_run_for_sha(
                'fake-token',
                'owner/repo',
                'feat/branch',
                'sync-nvl-stock.yml',
                'sha_123',
                require_event='workflow_dispatch',
                timeout_seconds=5,
            )
            self.assertEqual(run['id'], 102)

    def test_find_run_for_sha_rejects_wrong_sha_and_times_out(self):
        fake_runs = [
            {'id': 101, 'head_sha': 'old_sha_123', 'head_branch': 'feat/branch', 'name': 'Sync SharePoint NVL Stock'},
        ]
        with patch('requests.get') as mock_get:
            mock_get.return_value.status_code = 200
            mock_get.return_value.json.return_value = {'workflow_runs': fake_runs}

            with self.assertRaises(TimeoutError) as ctx:
                find_run_for_sha('fake-token', 'owner/repo', 'feat/branch', 'sync-nvl-stock.yml', 'new_sha_789', timeout_seconds=1, poll_interval=0.1)
            self.assertIn('new_sha_789', str(ctx.exception))
            self.assertIn('Hết thời gian', str(ctx.exception))

    def test_wait_for_run_completion_returns_when_completed(self):
        responses = [
            MagicMock(status_code=200, json=lambda: {'status': 'in_progress', 'conclusion': None, 'head_sha': 'sha1'}),
            MagicMock(status_code=200, json=lambda: {'status': 'completed', 'conclusion': 'success', 'head_sha': 'sha1'}),
        ]
        with patch('requests.get', side_effect=responses):
            result = wait_for_run_completion('fake-token', 'owner/repo', 102, poll_interval=0.1)
            self.assertEqual(result['status'], 'completed')
            self.assertEqual(result['conclusion'], 'success')

    def test_download_run_artifacts_extracts_zip_to_isolated_run_dir(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            base_dir = Path(tmpdir)
            buf = io.BytesIO()
            with zipfile.ZipFile(buf, 'w') as zf:
                zf.writestr('test_art.txt', 'hello artifact')
                zf.writestr('nvl_stock_report.json', json.dumps({'mode': 'dry_run'}))
            zip_bytes = buf.getvalue()

            mock_arts = {
                'artifacts': [
                    {'id': 555, 'name': 'nvl-audit-102', 'size_in_bytes': len(zip_bytes), 'archive_download_url': 'http://fake.url/zip'}
                ]
            }

            def fake_get(url, **kwargs):
                if 'artifacts' in url and not url.endswith('/zip'):
                    return MagicMock(status_code=200, json=lambda: mock_arts)
                return MagicMock(status_code=200, content=zip_bytes)

            with patch('requests.get', side_effect=fake_get):
                run_dir, files = download_run_artifacts('fake-token', 'owner/repo', 102, base_dir)
                self.assertEqual(run_dir, base_dir / 'run_102')
                self.assertIn('test_art.txt', files)
                self.assertIn('nvl_stock_report.json', files)
                self.assertTrue((run_dir / 'test_art.txt').is_file())
                self.assertTrue((run_dir / 'nvl_stock_report.json').is_file())

    def test_download_run_artifacts_fails_when_no_artifacts(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            base_dir = Path(tmpdir)
            with patch('requests.get') as mock_get:
                mock_get.return_value.status_code = 200
                mock_get.return_value.json.return_value = {'artifacts': []}
                with self.assertRaises(RuntimeError) as ctx:
                    download_run_artifacts('fake-token', 'owner/repo', 103, base_dir)
                self.assertIn('không có artifact nào', str(ctx.exception))

    def test_download_run_artifacts_fails_when_report_json_missing(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            base_dir = Path(tmpdir)
            buf = io.BytesIO()
            with zipfile.ZipFile(buf, 'w') as zf:
                zf.writestr('other.txt', 'no report here')
            zip_bytes = buf.getvalue()

            mock_arts = {
                'artifacts': [
                    {'id': 777, 'name': 'nvl-audit-104', 'size_in_bytes': len(zip_bytes), 'archive_download_url': 'http://fake.url/zip'}
                ]
            }

            def fake_get(url, **kwargs):
                if 'artifacts' in url and not url.endswith('/zip'):
                    return MagicMock(status_code=200, json=lambda: mock_arts)
                return MagicMock(status_code=200, content=zip_bytes)

            with patch('requests.get', side_effect=fake_get):
                with self.assertRaises(FileNotFoundError) as ctx:
                    download_run_artifacts('fake-token', 'owner/repo', 104, base_dir, require_report=True)
                self.assertIn('nvl_stock_report.json', str(ctx.exception))

    def test_download_run_artifacts_fails_on_corrupt_zip(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            base_dir = Path(tmpdir)
            mock_arts = {
                'artifacts': [
                    {'id': 888, 'name': 'corrupted-art', 'size_in_bytes': 100, 'archive_download_url': 'http://fake.url/zip'}
                ]
            }

            def fake_get(url, **kwargs):
                if 'artifacts' in url and not url.endswith('/zip'):
                    return MagicMock(status_code=200, json=lambda: mock_arts)
                return MagicMock(status_code=200, content=b"THIS IS NOT A ZIP FILE")

            with patch('requests.get', side_effect=fake_get):
                with self.assertRaises(RuntimeError) as ctx:
                    download_run_artifacts('fake-token', 'owner/repo', 105, base_dir)
                self.assertIn('không phải file zip hợp lệ', str(ctx.exception))

    def test_verify_run_execution_success(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            p = Path(tmpdir)
            cfg_file = p / 'test_cfg.json'
            cfg_file.write_text(
                json.dumps({
                    'source': {
                        'sharepoint_path': 'Tinh san xuat Mua hang 2027/Ton He thong/Ton Ke Toan/XNT_ketoan_Vikoda.xlsm',
                        'name': 'XNT_ketoan_Vikoda.xlsm',
                    },
                    'target': {
                        'sharepoint_path': 'Tinh san xuat Mua hang 2027/Kế hoạch mua hàng.xlsx',
                        'name': 'Kế hoạch mua hàng.xlsx',
                    }
                }),
                encoding='utf-8',
            )

            rep_file = p / 'nvl_stock_report.json'
            rep_file.write_text(
                json.dumps({
                    'mode': 'dry_run',
                    'status': 'completed_with_warnings',
                    'source': {
                        'sharepoint_path': 'Tinh san xuat Mua hang 2027/Ton He thong/Ton Ke Toan/XNT_ketoan_Vikoda.xlsm',
                        'name': 'XNT_ketoan_Vikoda.xlsm',
                    },
                    'target': {
                        'sharepoint_path': 'Tinh san xuat Mua hang 2027/Kế hoạch mua hàng.xlsx',
                        'name': 'Kế hoạch mua hàng.xlsx',
                    },
                }),
                encoding='utf-8',
            )

            audit_file = p / 'audit_summary.json'
            audit_file.write_text(
                json.dumps({'config_file': str(cfg_file)}),
                encoding='utf-8',
            )

            # verify dry-run matches expected_publish=False
            verify_run_execution(p, str(cfg_file), expected_publish=False)

    def test_verify_run_execution_rejects_nonexistent_config(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            p = Path(tmpdir)
            with self.assertRaises(FileNotFoundError) as ctx:
                verify_run_execution(p, 'non_existent_config.json', expected_publish=False)
            self.assertIn('không tồn tại', str(ctx.exception))

    def test_verify_run_execution_rejects_report_missing_target_info(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            p = Path(tmpdir)
            cfg_file = p / 'valid_cfg.json'
            cfg_file.write_text(json.dumps({'target': {'sharepoint_path': 'a', 'name': 'b'}}), encoding='utf-8')

            rep_file = p / 'nvl_stock_report.json'
            # Report thiếu hoàn toàn trường 'target'
            rep_file.write_text(
                json.dumps({
                    'mode': 'dry_run',
                    'source': {'name': 'src', 'sharepoint_path': 'p_src'},
                }),
                encoding='utf-8',
            )

            with self.assertRaises(ValueError) as ctx:
                verify_run_execution(p, str(cfg_file), expected_publish=False)
            self.assertIn("Report thiếu thông tin bắt buộc về 'source' hoặc 'target'", str(ctx.exception))

            # Report có 'target' nhưng thiếu 'sharepoint_path'
            rep_file.write_text(
                json.dumps({
                    'mode': 'dry_run',
                    'source': {'name': 'src', 'sharepoint_path': 'p_src'},
                    'target': {'name': 'tgt'},
                }),
                encoding='utf-8',
            )
            with self.assertRaises(ValueError) as ctx:
                verify_run_execution(p, str(cfg_file), expected_publish=False)
            self.assertIn("Report thiếu trường bắt buộc của đích", str(ctx.exception))

    def test_verify_run_execution_rejects_failed_status_report(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            p = Path(tmpdir)
            cfg_file = p / 'valid_cfg.json'
            cfg_file.write_text(json.dumps({'target': {}}), encoding='utf-8')

            rep_file = p / 'nvl_stock_report.json'
            rep_file.write_text(
                json.dumps({
                    'mode': 'failed',
                    'status': 'failed',
                    'phase': 'fetch_target',
                    'error_message': 'ItemNotFound 404',
                }),
                encoding='utf-8',
            )

            with self.assertRaises(ValueError) as ctx:
                verify_run_execution(p, str(cfg_file), expected_publish=False)
            self.assertIn('Report ghi nhận trạng thái thất bại', str(ctx.exception))

    def test_verify_run_execution_rejects_publish_when_expecting_dryrun(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            p = Path(tmpdir)
            cfg_file = p / 'valid_cfg.json'
            cfg_file.write_text(json.dumps({'target': {}}), encoding='utf-8')

            rep_file = p / 'nvl_stock_report.json'
            rep_file.write_text(json.dumps({'mode': 'publish'}), encoding='utf-8')

            with self.assertRaises(ValueError) as ctx:
                verify_run_execution(p, str(cfg_file), expected_publish=False)
            self.assertIn('nguy cơ ghi đè', str(ctx.exception))

    def test_verify_run_execution_rejects_dryrun_when_expecting_publish(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            p = Path(tmpdir)
            cfg_file = p / 'valid_cfg.json'
            cfg_file.write_text(json.dumps({'target': {}}), encoding='utf-8')

            rep_file = p / 'nvl_stock_report.json'
            rep_file.write_text(json.dumps({'mode': 'dry_run'}), encoding='utf-8')

            with self.assertRaises(ValueError) as ctx:
                verify_run_execution(p, str(cfg_file), expected_publish=True)
            self.assertIn('yêu cầu publish=True nhưng report ghi nhận mode=\'dry_run\'', str(ctx.exception))

    def test_verify_run_execution_rejects_target_path_mismatch(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            p = Path(tmpdir)
            cfg_file = p / 'staging_cfg.json'
            cfg_file.write_text(
                json.dumps({
                    'source': {
                        'sharepoint_path': 'Tinh san xuat Mua hang 2027/Ton He thong/Ton Ke Toan/XNT_ketoan_Vikoda.xlsm',
                        'name': 'XNT_ketoan_Vikoda.xlsm',
                    },
                    'target': {
                        'sharepoint_path': 'Tinh san xuat Mua hang 2027/Test_Ke_hoach_mua_hang_copy.xlsx',
                        'name': 'Test_Ke_hoach_mua_hang_copy.xlsx',
                    }
                }),
                encoding='utf-8',
            )

            rep_file = p / 'nvl_stock_report.json'
            rep_file.write_text(
                json.dumps({
                    'mode': 'dry_run',
                    'source': {
                        'sharepoint_path': 'Tinh san xuat Mua hang 2027/Ton He thong/Ton Ke Toan/XNT_ketoan_Vikoda.xlsm',
                        'name': 'XNT_ketoan_Vikoda.xlsm',
                    },
                    'target': {
                        'sharepoint_path': 'Tinh san xuat Mua hang 2027/Kế hoạch mua hàng.xlsx',
                        'name': 'Kế hoạch mua hàng.xlsx',
                    },
                }),
                encoding='utf-8',
            )

            with self.assertRaises(ValueError) as ctx:
                verify_run_execution(p, str(cfg_file), expected_publish=False)
            self.assertIn('Cấu hình đường dẫn đích không khớp', str(ctx.exception))

    def test_verify_run_execution_rejects_audit_config_file_mismatch(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            p = Path(tmpdir)
            cfg_file = p / 'staging_cfg.json'
            cfg_file.write_text(json.dumps({'target': {}}), encoding='utf-8')

            rep_file = p / 'nvl_stock_report.json'
            rep_file.write_text(
                json.dumps({
                    'mode': 'dry_run',
                    'source': {'name': 'src', 'sharepoint_path': 'p_src'},
                    'target': {'name': 'tgt', 'sharepoint_path': 'p_tgt'},
                }),
                encoding='utf-8',
            )

            audit_file = p / 'audit_summary.json'
            audit_file.write_text(
                json.dumps({'config_file': 'some_other_config.json'}),
                encoding='utf-8',
            )

            with self.assertRaises(ValueError) as ctx:
                verify_run_execution(p, str(cfg_file), expected_publish=False)
            self.assertIn('Config file trong audit_summary', str(ctx.exception))

    def test_main_fails_when_conclusion_is_failure(self):
        with patch('scripts.trigger_and_download_dryrun.get_token', return_value='fake-token'):
            with patch('scripts.trigger_and_download_dryrun.dispatch_workflow'):
                with patch('scripts.trigger_and_download_dryrun.find_run_for_sha', return_value={'id': 999, 'head_sha': 'sha_fail'}):
                    with patch('scripts.trigger_and_download_dryrun.wait_for_run_completion', return_value={'status': 'completed', 'conclusion': 'failure', 'head_sha': 'sha_fail'}):
                        with patch('scripts.trigger_and_download_dryrun.download_run_artifacts', return_value=(Path('/tmp'), [])):
                            exit_code = main(['--sha', 'sha_fail', '--no-dispatch'])
                            self.assertEqual(exit_code, 1)

    def test_main_succeeds_when_conclusion_is_success(self):
        with patch('scripts.trigger_and_download_dryrun.get_token', return_value='fake-token'):
            with patch('scripts.trigger_and_download_dryrun.dispatch_workflow'):
                with patch('scripts.trigger_and_download_dryrun.find_run_for_sha', return_value={'id': 888, 'head_sha': 'sha_ok'}):
                    with patch('scripts.trigger_and_download_dryrun.wait_for_run_completion', return_value={'status': 'completed', 'conclusion': 'success', 'head_sha': 'sha_ok'}):
                        with patch('scripts.trigger_and_download_dryrun.download_run_artifacts', return_value=(Path('/tmp'), [])):
                            with patch('scripts.trigger_and_download_dryrun.verify_run_execution'):
                                exit_code = main(['--sha', 'sha_ok', '--no-dispatch'])
                                self.assertEqual(exit_code, 0)


if __name__ == '__main__':
    unittest.main()
