"""Immutable release manifest for successful NVL production publishes."""

from __future__ import annotations

import argparse
from io import BytesIO
import hashlib
import json
import os
from datetime import datetime, timezone
from pathlib import Path
from typing import Any
from zipfile import BadZipFile, ZipFile


NVL_RELEASE_MANIFEST_FILE = Path("nvl_release_manifest.json")
NVL_READINESS_REPORT_FILE = Path("nvl_production_readiness_report.json")
NVL_RELEASE_SCHEMA = "nvl_release_manifest_v1"
NVL_RELEASE_SCHEMA_VERSION = 1
VOLATILE_XLSX_PARTS = {"docProps/core.xml"}


def _sha256_bytes(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def stable_workbook_sha256(data: bytes) -> str:
    """Hash XLSX payload while ignoring volatile Office metadata."""
    raw_sha256 = _sha256_bytes(data)
    try:
        with ZipFile(BytesIO(data), "r") as archive:
            names = sorted(
                name
                for name in archive.namelist()
                if name not in VOLATILE_XLSX_PARTS
            )
            if not names:
                return raw_sha256

            digest = hashlib.sha256()
            for name in names:
                name_bytes = name.encode("utf-8")
                payload = archive.read(name)
                digest.update(len(name_bytes).to_bytes(4, "big"))
                digest.update(name_bytes)
                digest.update(len(payload).to_bytes(8, "big"))
                digest.update(payload)
            return digest.hexdigest()
    except (BadZipFile, OSError):
        return raw_sha256


def _read_json(path: Path, *, label: str) -> dict[str, Any]:
    try:
        value = json.loads(path.read_text(encoding="utf-8"))
    except Exception as exc:
        raise RuntimeError(f"Không đọc được {label}: {path}") from exc
    if not isinstance(value, dict):
        raise RuntimeError(f"{label} phải là JSON object: {path}")
    return value


def _read_readiness_report(
    path: Path = NVL_READINESS_REPORT_FILE,
    *,
    required: bool,
) -> dict[str, Any]:
    if not path.exists():
        if required:
            raise RuntimeError(
                "NVL production publish thiếu nvl_production_readiness_report.json."
            )
        return {
            "status": "not_available",
            "gate_version": None,
            "git_sha": None,
        }

    readiness = _read_json(path, label="NVL readiness report")
    if required and readiness.get("status") != "passed":
        raise RuntimeError(
            "NVL production readiness gate chưa PASS; không tạo release manifest."
        )
    return readiness


def _require_successful_publish(
    stock_report: dict[str, Any],
    open_po_report: dict[str, Any],
) -> None:
    if stock_report.get("mode") != "publish":
        raise RuntimeError(
            "NVL stock report không phải production publish."
        )
    if stock_report.get("status") not in {
        "published",
        "published_with_warnings",
        "unchanged",
    }:
        raise RuntimeError(
            "NVL stock publish chưa hoàn tất thành công."
        )
    if not open_po_report.get("publish_requested"):
        raise RuntimeError(
            "Open-PO report không ghi nhận publish_requested=true."
        )
    if open_po_report.get("published") is not True:
        raise RuntimeError(
            "Open-PO publish chưa hoàn tất thành công."
        )


def build_release_manifest(
    stock_report: dict[str, Any],
    open_po_report: dict[str, Any],
    stock_proposal_bytes: bytes,
    final_workbook_bytes: bytes,
    *,
    published_at: str | None = None,
    environ: dict[str, str] | None = None,
    readiness_path: Path = NVL_READINESS_REPORT_FILE,
) -> dict[str, Any]:
    """Build one release record for the final Ton_NVL state after D + E publish."""
    _require_successful_publish(stock_report, open_po_report)

    env = dict(os.environ if environ is None else environ)
    strict = str(env.get("NVL_REQUIRE_READINESS", "")).strip() in {
        "1",
        "true",
        "TRUE",
        "yes",
        "YES",
    }
    readiness = _read_readiness_report(
        readiness_path,
        required=strict,
    )

    commit_sha = (
        str(env.get("GITHUB_SHA") or "").strip()
        or str(readiness.get("git_sha") or "").strip()
    )
    readiness_sha = str(readiness.get("git_sha") or "").strip()

    if strict and not commit_sha:
        raise RuntimeError(
            "NVL production release thiếu commit SHA."
        )
    if strict and readiness_sha and readiness_sha != commit_sha:
        raise RuntimeError(
            "Commit SHA của NVL readiness report không khớp release commit."
        )

    stock_source = stock_report.get("source") or {}
    stock_target = stock_report.get("target") or {}
    stock_upload = stock_report.get("upload_result") or {}
    open_source = open_po_report.get("source") or {}
    open_target = open_po_report.get("target") or {}
    open_evidence = open_po_report.get("publish_evidence") or {}

    stock_raw = _sha256_bytes(stock_proposal_bytes)
    stock_stable = stable_workbook_sha256(stock_proposal_bytes)
    final_raw = _sha256_bytes(final_workbook_bytes)
    final_stable = stable_workbook_sha256(final_workbook_bytes)

    run_id = str(env.get("GITHUB_RUN_ID") or "").strip()
    run_attempt = str(env.get("GITHUB_RUN_ATTEMPT") or "1").strip()
    timestamp = published_at or datetime.now(timezone.utc).isoformat()
    instance = (
        f"gh-{run_id}-{run_attempt}"
        if run_id
        else timestamp.replace(":", "").replace("+", "_")
    )
    release_id = f"nvl-{final_stable[:16]}-{instance}"
    release_version = (
        f"{NVL_RELEASE_SCHEMA}.workbook-{final_stable[:16]}"
        + (f".run-{run_id}.{run_attempt}" if run_id else "")
    )

    manifest = {
        "schema": NVL_RELEASE_SCHEMA,
        "schema_version": NVL_RELEASE_SCHEMA_VERSION,
        "release_id": release_id,
        "release_version": release_version,
        "published_at": timestamp,
        "commit": {
            "sha": commit_sha or None,
            "ref": env.get("GITHUB_REF") or None,
            "repository": env.get("GITHUB_REPOSITORY") or None,
            "run_id": run_id or None,
            "run_attempt": run_attempt if run_id else None,
            "workflow": env.get("GITHUB_WORKFLOW") or None,
            "actor": env.get("GITHUB_ACTOR") or None,
        },
        "readiness": {
            "status": readiness.get("status"),
            "gate_version": readiness.get("gate_version"),
            "git_sha": readiness.get("git_sha"),
            "report_sha256": (
                _sha256_bytes(readiness_path.read_bytes())
                if readiness_path.exists()
                else None
            ),
        },
        "input_revision": {
            "stock_source_etag": stock_source.get("revision"),
            "stock_target_etag_before": stock_target.get("revision"),
            "stock_target_etag_after": (
                stock_upload.get("eTag")
                or (
                    stock_target.get("revision")
                    if stock_report.get("status") == "unchanged"
                    else None
                )
            ),
            "open_po_source_etag": open_source.get("revision"),
            "open_po_target_etag_before": open_target.get("revision_before"),
            "open_po_target_etag_after": (
                open_evidence.get("target_revision_after")
                or open_target.get("revision_after")
            ),
        },
        "artifacts": {
            "stock_proposal": {
                "sha256": stock_raw,
                "stable_sha256": stock_stable,
            },
            "open_po_proposal": {
                "sha256": final_raw,
                "stable_sha256": final_stable,
            },
        },
        "published_workbook": {
            "sha256": open_evidence.get("server_sha256") or final_raw,
            "proposal_sha256": final_raw,
            "stable_sha256": final_stable,
            "post_upload_verified": open_evidence.get(
                "post_upload_verified"
            ),
        },
        "publish_evidence": {
            "stock": {
                "status": stock_report.get("status"),
                "changed_count": (
                    (stock_report.get("metrics") or {}).get("changed_count")
                ),
                "post_upload_verified": stock_report.get(
                    "post_upload_verified",
                    stock_report.get("status") == "unchanged",
                ),
                "upload_result": stock_report.get("upload_result"),
            },
            "open_po": {
                "published": open_po_report.get("published"),
                "changed_cells": open_target.get("changed_cells"),
                "evidence": open_evidence,
            },
        },
    }
    return json.loads(
        json.dumps(manifest, ensure_ascii=False, default=str)
    )


def write_release_manifest(
    manifest: dict[str, Any],
    *,
    path: Path = NVL_RELEASE_MANIFEST_FILE,
) -> dict[str, Any]:
    path.write_text(
        json.dumps(
            manifest,
            ensure_ascii=False,
            indent=2,
            sort_keys=True,
        )
        + "\n",
        encoding="utf-8",
    )
    return manifest


def save_release_manifest(
    stock_report: dict[str, Any],
    open_po_report: dict[str, Any],
    stock_proposal_bytes: bytes,
    final_workbook_bytes: bytes,
    *,
    path: Path = NVL_RELEASE_MANIFEST_FILE,
    readiness_path: Path = NVL_READINESS_REPORT_FILE,
) -> dict[str, Any]:
    manifest = build_release_manifest(
        stock_report,
        open_po_report,
        stock_proposal_bytes,
        final_workbook_bytes,
        readiness_path=readiness_path,
    )
    return write_release_manifest(manifest, path=path)


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="Ghi NVL release manifest sau production publish thành công."
    )
    parser.add_argument(
        "--stock-report",
        default="nvl_stock_report.json",
    )
    parser.add_argument(
        "--open-po-report",
        default="nvl_open_po_report.json",
    )
    parser.add_argument(
        "--stock-proposal",
        default="nvl_stock_proposal.xlsx",
    )
    parser.add_argument(
        "--final-workbook",
        default="nvl_open_po_proposal.xlsx",
    )
    parser.add_argument(
        "--readiness-report",
        default=str(NVL_READINESS_REPORT_FILE),
    )
    parser.add_argument(
        "--out",
        default=str(NVL_RELEASE_MANIFEST_FILE),
    )
    args = parser.parse_args(argv)

    stock_report_path = Path(args.stock_report)
    open_po_report_path = Path(args.open_po_report)
    stock_proposal_path = Path(args.stock_proposal)
    final_workbook_path = Path(args.final_workbook)

    manifest = build_release_manifest(
        _read_json(stock_report_path, label="NVL stock report"),
        _read_json(open_po_report_path, label="NVL open-PO report"),
        stock_proposal_path.read_bytes(),
        final_workbook_path.read_bytes(),
        readiness_path=Path(args.readiness_report),
    )
    write_release_manifest(manifest, path=Path(args.out))
    print(
        f"[NVL-RELEASE] release_id={manifest['release_id']} "
        f"manifest={args.out}"
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
