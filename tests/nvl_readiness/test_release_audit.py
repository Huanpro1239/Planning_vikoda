"""Contracts for NVL release ledger audit and index."""

from __future__ import annotations

import hashlib
import json
from pathlib import Path
import shutil
import subprocess
import sys
import tempfile
import unittest

from nvl.release import (
    build_release_manifest,
    write_release_manifest,
)
from nvl.release_audit import audit_nvl_release_ledger


ROOT = Path(__file__).resolve().parents[2]


class NVLReleaseAuditTests(unittest.TestCase):
    def _reports(self, *, suffix: str):
        stock = {
            "mode": "publish",
            "status": "published",
            "source": {
                "revision": f'"stock-source-{suffix}"',
            },
            "target": {
                "revision": f'"target-before-stock-{suffix}"',
            },
            "metrics": {
                "changed_count": 2,
            },
            "post_upload_verified": True,
            "upload_result": {
                "name": "Kế hoạch mua hàng.xlsx",
                "eTag": f'"target-after-stock-{suffix}"',
            },
        }
        final_bytes = f"final-workbook-{suffix}".encode("utf-8")
        open_po = {
            "publish_requested": True,
            "published": True,
            "source": {
                "revision": f'"open-po-source-{suffix}"',
            },
            "target": {
                "revision_before": f'"target-before-open-po-{suffix}"',
                "revision_after": f'"target-after-open-po-{suffix}"',
                "changed_cells": 3,
            },
            "publish_evidence": {
                "expected_etag": f'"target-before-open-po-{suffix}"',
                "upload_skipped": False,
                "upload_etag": f'"target-after-open-po-{suffix}"',
                "target_revision_after": f'"target-after-open-po-{suffix}"',
                "post_upload_verified": True,
                "server_sha256": hashlib.sha256(final_bytes).hexdigest(),
            },
        }
        return stock, open_po, final_bytes

    def _readiness(self, root: Path):
        path = root / "nvl_production_readiness_report.json"
        path.write_text(
            json.dumps(
                {
                    "status": "passed",
                    "gate_version": "nvl_production_readiness_v1",
                    "git_sha": "deadbeef",
                },
                sort_keys=True,
            ),
            encoding="utf-8",
        )
        return path

    def _append_release(
        self,
        root: Path,
        releases: Path,
        *,
        suffix: str,
        timestamp: str,
        run_id: str,
        previous_path: Path | None = None,
    ) -> Path:
        stock, open_po, final_bytes = self._reports(suffix=suffix)
        stock_bytes = f"stock-proposal-{suffix}".encode("utf-8")
        readiness = self._readiness(root)

        previous_manifest = None
        previous_sha = None
        if previous_path is not None:
            previous_raw = previous_path.read_bytes()
            previous_sha = hashlib.sha256(previous_raw).hexdigest()
            previous_manifest = json.loads(
                previous_raw.decode("utf-8")
            )

        manifest = build_release_manifest(
            stock,
            open_po,
            stock_bytes,
            final_bytes,
            published_at=timestamp,
            environ={
                "NVL_REQUIRE_READINESS": "1",
                "NVL_RELEASE_COMMIT_SHA": "deadbeef",
                "GITHUB_SHA": "workflow-context-sha",
                "GITHUB_RUN_ID": run_id,
                "GITHUB_RUN_ATTEMPT": "1",
                "NVL_UPSTREAM_PLANNING_WORKFLOW": "Sync SharePoint Stock",
                "NVL_UPSTREAM_PLANNING_RUN_ID": f"planning-{run_id}",
                "NVL_UPSTREAM_PLANNING_HEAD_SHA": "deadbeef",
                "NVL_UPSTREAM_PLANNING_EVENT": "schedule",
            },
            readiness_path=readiness,
            previous_manifest=previous_manifest,
            previous_manifest_sha256=previous_sha,
        )
        path = releases / f"{manifest['release_id']}.json"
        write_release_manifest(manifest, path=path)
        return path

    def _healthy_ledger(self, root: Path):
        nvl = root / "nvl"
        releases = nvl / "releases"
        releases.mkdir(parents=True)
        first = self._append_release(
            root,
            releases,
            suffix="one",
            timestamp="2026-09-25T01:00:00+00:00",
            run_id="100",
        )
        second = self._append_release(
            root,
            releases,
            suffix="two",
            timestamp="2026-09-25T02:00:00+00:00",
            run_id="101",
            previous_path=first,
        )
        latest = nvl / "latest_release.json"
        shutil.copyfile(second, latest)
        return releases, latest, first, second

    def test_healthy_hash_chain_and_latest_pointer_pass(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            root = Path(tmpdir)
            releases, latest, _, _ = self._healthy_ledger(root)
            index = audit_nvl_release_ledger(
                releases,
                latest_path=latest,
                verify_commit=False,
            )

        self.assertEqual(index["status"], "passed")
        self.assertEqual(index["summary"]["release_count"], 2)
        self.assertEqual(index["latest"]["status"], "passed")
        self.assertEqual(
            [row["chain_status"] for row in index["releases"]],
            ["anchor", "linked"],
        )
        self.assertEqual(
            index["releases"][0]["upstream_planning_run_id"],
            "planning-100",
        )
        self.assertEqual(
            index["releases"][1]["upstream_planning_head_sha"],
            "deadbeef",
        )
        self.assertEqual(
            index["releases"][1]["upstream_planning_event"],
            "schedule",
        )

    def test_missing_predecessor_is_detected(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            root = Path(tmpdir)
            releases, latest, first, _ = self._healthy_ledger(root)
            first.unlink()
            index = audit_nvl_release_ledger(
                releases,
                latest_path=latest,
                verify_commit=False,
            )

        self.assertEqual(index["status"], "failed")
        codes = {
            issue["code"]
            for row in index["releases"]
            for issue in row["issues"]
        }
        self.assertIn("missing_predecessor", codes)

    def test_tampered_predecessor_breaks_hash_chain(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            root = Path(tmpdir)
            releases, latest, first, _ = self._healthy_ledger(root)
            manifest = json.loads(first.read_text(encoding="utf-8"))
            manifest["commit"]["actor"] = "tampered-actor"
            first.write_text(
                json.dumps(
                    manifest,
                    ensure_ascii=False,
                    indent=2,
                    sort_keys=True,
                )
                + "\n",
                encoding="utf-8",
            )
            index = audit_nvl_release_ledger(
                releases,
                latest_path=latest,
                verify_commit=False,
            )

        self.assertEqual(index["status"], "failed")
        codes = {
            issue["code"]
            for row in index["releases"]
            for issue in row["issues"]
        }
        self.assertIn("predecessor_hash_mismatch", codes)

    def test_wrong_latest_pointer_is_detected(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            root = Path(tmpdir)
            releases, latest, first, _ = self._healthy_ledger(root)
            shutil.copyfile(first, latest)
            index = audit_nvl_release_ledger(
                releases,
                latest_path=latest,
                verify_commit=False,
            )

        self.assertEqual(index["status"], "failed")
        self.assertEqual(index["latest"]["status"], "failed")

    def test_chain_fork_or_gap_is_detected(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            root = Path(tmpdir)
            releases, latest, first, second = self._healthy_ledger(root)
            third = self._append_release(
                root,
                releases,
                suffix="three",
                timestamp="2026-09-25T03:00:00+00:00",
                run_id="102",
                previous_path=first,
            )
            shutil.copyfile(third, latest)
            index = audit_nvl_release_ledger(
                releases,
                latest_path=latest,
                verify_commit=False,
            )

        self.assertEqual(index["status"], "failed")
        codes = {
            issue["code"]
            for row in index["releases"]
            for issue in row["issues"]
        }
        self.assertIn("chain_gap_or_branch", codes)
        self.assertIn("chain_fork", codes)

    def test_filename_release_id_mismatch_is_detected(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            root = Path(tmpdir)
            releases, latest, first, _ = self._healthy_ledger(root)
            renamed = releases / "wrong-name.json"
            first.rename(renamed)
            index = audit_nvl_release_ledger(
                releases,
                latest_path=latest,
                verify_commit=False,
            )

        self.assertEqual(index["status"], "failed")
        codes = {
            issue["code"]
            for row in index["releases"]
            for issue in row["issues"]
        }
        self.assertIn("filename_release_id_mismatch", codes)

    def test_legacy_unlinked_manifest_is_warning_not_tamper(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            root = Path(tmpdir)
            nvl = root / "nvl"
            releases = nvl / "releases"
            releases.mkdir(parents=True)
            release = self._append_release(
                root,
                releases,
                suffix="legacy",
                timestamp="2026-09-25T01:00:00+00:00",
                run_id="90",
            )
            manifest = json.loads(release.read_text(encoding="utf-8"))
            manifest.pop("chain", None)
            release.write_text(
                json.dumps(
                    manifest,
                    ensure_ascii=False,
                    indent=2,
                    sort_keys=True,
                )
                + "\n",
                encoding="utf-8",
            )
            latest = nvl / "latest_release.json"
            shutil.copyfile(release, latest)

            index = audit_nvl_release_ledger(
                releases,
                latest_path=latest,
                verify_commit=False,
            )

        self.assertEqual(index["status"], "passed")
        self.assertEqual(
            index["summary"]["legacy_unlinked_count"],
            1,
        )
        self.assertEqual(
            index["releases"][0]["chain_status"],
            "legacy_unlinked",
        )
        codes = {
            issue["code"]
            for issue in index["releases"][0]["issues"]
        }
        self.assertIn("legacy_unlinked", codes)

    def test_empty_ledger_is_not_a_failure(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            releases = Path(tmpdir) / "nvl" / "releases"
            index = audit_nvl_release_ledger(
                releases,
                verify_commit=False,
            )

        self.assertEqual(index["status"], "empty")
        self.assertEqual(index["summary"]["release_count"], 0)
        self.assertEqual(index["latest"]["status"], "not_applicable")

    def test_cli_writes_json_and_csv_index(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            root = Path(tmpdir)
            releases, latest, _, _ = self._healthy_ledger(root)
            out_json = root / "index.json"
            out_csv = root / "index.csv"
            result = subprocess.run(
                [
                    sys.executable,
                    "-X",
                    "utf8",
                    str(ROOT / "scripts" / "audit_nvl_releases.py"),
                    str(releases),
                    "--latest",
                    str(latest),
                    "--out-json",
                    str(out_json),
                    "--out-csv",
                    str(out_csv),
                    "--no-git",
                ],
                cwd=ROOT,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True,
                encoding="utf-8",
                errors="replace",
                check=False,
            )

            self.assertEqual(
                result.returncode,
                0,
                msg=result.stderr or result.stdout,
            )
            self.assertTrue(out_json.is_file())
            self.assertTrue(out_csv.is_file())
            payload = json.loads(out_json.read_text(encoding="utf-8"))
            self.assertEqual(payload["status"], "passed")
            csv_text = out_csv.read_text(encoding="utf-8-sig")
            self.assertIn("release_id", csv_text)
            self.assertIn("chain_status", csv_text)
            self.assertIn("upstream_planning_run_id", csv_text)
            self.assertIn("upstream_planning_head_sha", csv_text)


if __name__ == "__main__":
    unittest.main()
