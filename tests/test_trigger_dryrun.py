"""Unit tests for scripts/trigger_and_download_dryrun.py."""

import io
import tempfile
import unittest
import zipfile
from pathlib import Path
from unittest.mock import MagicMock, patch

from scripts.trigger_and_download_dryrun import (

    dispatch_workflow,
    download_run_artifacts,
    find_run_for_sha,
    main,
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
            self.assertIn('Tuyệt đối không nhận run cũ', str(ctx.exception))

    def test_wait_for_run_completion_returns_when_completed(self):
        responses = [
            MagicMock(status_code=200, json=lambda: {'status': 'in_progress', 'conclusion': None, 'head_sha': 'sha1'}),
            MagicMock(status_code=200, json=lambda: {'status': 'completed', 'conclusion': 'success', 'head_sha': 'sha1'}),
        ]
        with patch('requests.get', side_effect=responses):
            result = wait_for_run_completion('fake-token', 'owner/repo', 102, poll_interval=0.1)
            self.assertEqual(result['status'], 'completed')
            self.assertEqual(result['conclusion'], 'success')

    def test_download_run_artifacts_extracts_zip(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            out_dir = Path(tmpdir)
            buf = io.BytesIO()
            with zipfile.ZipFile(buf, 'w') as zf:
                zf.writestr('test_art.txt', 'hello artifact')
            zip_bytes = buf.getvalue()

            mock_arts = {
                'artifacts': [
                    {'name': 'audit-102', 'size_in_bytes': len(zip_bytes), 'archive_download_url': 'http://fake.url/zip'}
                ]
            }

            def fake_get(url, **kwargs):
                if 'artifacts' in url and not url.endswith('/zip'):
                    return MagicMock(status_code=200, json=lambda: mock_arts)
                return MagicMock(status_code=200, content=zip_bytes)

            with patch('requests.get', side_effect=fake_get):
                files = download_run_artifacts('fake-token', 'owner/repo', 102, out_dir, timeout_seconds=2)
                self.assertIn('test_art.txt', files)
                self.assertTrue((out_dir / 'test_art.txt').is_file())

    def test_main_fails_when_conclusion_is_failure(self):
        with patch('scripts.trigger_and_download_dryrun.get_token', return_value='fake-token'):
            with patch('scripts.trigger_and_download_dryrun.dispatch_workflow'):
                with patch('scripts.trigger_and_download_dryrun.find_run_for_sha', return_value={'id': 999, 'head_sha': 'sha_fail'}):
                    with patch('scripts.trigger_and_download_dryrun.wait_for_run_completion', return_value={'status': 'completed', 'conclusion': 'failure', 'head_sha': 'sha_fail'}):
                        with patch('scripts.trigger_and_download_dryrun.download_run_artifacts'):
                            exit_code = main(['--sha', 'sha_fail', '--no-dispatch'])
                            self.assertEqual(exit_code, 1)

    def test_main_succeeds_when_conclusion_is_success(self):
        with patch('scripts.trigger_and_download_dryrun.get_token', return_value='fake-token'):
            with patch('scripts.trigger_and_download_dryrun.dispatch_workflow'):
                with patch('scripts.trigger_and_download_dryrun.find_run_for_sha', return_value={'id': 888, 'head_sha': 'sha_ok'}):
                    with patch('scripts.trigger_and_download_dryrun.wait_for_run_completion', return_value={'status': 'completed', 'conclusion': 'success', 'head_sha': 'sha_ok'}):
                        with patch('scripts.trigger_and_download_dryrun.download_run_artifacts'):
                            exit_code = main(['--sha', 'sha_ok', '--no-dispatch'])
                            self.assertEqual(exit_code, 0)


if __name__ == '__main__':
    unittest.main()
