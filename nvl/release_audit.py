"""Audit and index the immutable NVL release ledger."""

from __future__ import annotations

import csv
import hashlib
import json
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

from nvl.verify_release import (
    NVLReleaseVerificationError,
    verify_nvl_release,
)


INDEX_SCHEMA = "nvl_release_index_v1"
INDEX_SCHEMA_VERSION = 1


def _sha256(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def _load_json_bytes(data: bytes, *, label: str) -> dict[str, Any]:
    try:
        value = json.loads(data.decode("utf-8"))
    except Exception as exc:
        raise ValueError(f"Không đọc được {label}") from exc
    if not isinstance(value, dict):
        raise ValueError(f"{label} phải là JSON object")
    return value


def _parse_timestamp(value: Any) -> datetime | None:
    text = str(value or "").strip()
    if not text:
        return None
    try:
        parsed = datetime.fromisoformat(text.replace("Z", "+00:00"))
        if parsed.tzinfo is None:
            parsed = parsed.replace(tzinfo=timezone.utc)
        return parsed.astimezone(timezone.utc)
    except ValueError:
        return None


def _issue(
    record: dict[str, Any],
    code: str,
    detail: str,
    *,
    severity: str = "error",
) -> None:
    record["issues"].append(
        {
            "severity": severity,
            "code": code,
            "detail": detail,
        }
    )


def _record_status(record: dict[str, Any]) -> str:
    if any(
        issue.get("severity") == "error"
        for issue in record.get("issues", [])
    ):
        return "failed"
    if any(
        issue.get("severity") == "warning"
        for issue in record.get("issues", [])
    ):
        return "warning"
    return "passed"


def audit_nvl_release_ledger(
    releases_dir: str | Path,
    *,
    latest_path: str | Path | None = None,
    verify_commit: bool = True,
    repo_root: str | Path | None = None,
) -> dict[str, Any]:
    """Verify every ledger manifest and audit hash-chain continuity."""
    releases_dir = Path(releases_dir)
    if latest_path is None:
        latest_path = releases_dir.parent / "latest_release.json"
    else:
        latest_path = Path(latest_path)

    manifest_paths = (
        sorted(releases_dir.glob("*.json"))
        if releases_dir.is_dir()
        else []
    )
    records: list[dict[str, Any]] = []

    for path in manifest_paths:
        raw = path.read_bytes()
        record: dict[str, Any] = {
            "path": str(path),
            "filename": path.name,
            "manifest_sha256": _sha256(raw),
            "release_id": None,
            "published_at": None,
            "commit_sha": None,
            "gate_version": None,
            "published_workbook_sha256": None,
            "published_workbook_stable_sha256": None,
            "stock_source_etag": None,
            "open_po_source_etag": None,
            "previous_release_id": None,
            "previous_manifest_sha256": None,
            "chain_status": "unknown",
            "verification_status": "unknown",
            "verification_warning_count": 0,
            "issues": [],
            "_manifest": None,
            "_published_dt": None,
        }

        try:
            manifest = _load_json_bytes(
                raw,
                label=f"manifest {path.name}",
            )
            record["_manifest"] = manifest
        except ValueError as exc:
            _issue(record, "invalid_json", str(exc))
            record["verification_status"] = "failed"
            record["chain_status"] = "invalid"
            record["status"] = _record_status(record)
            records.append(record)
            continue

        release_id = str(manifest.get("release_id") or "").strip()
        published_at = str(manifest.get("published_at") or "").strip()
        commit = manifest.get("commit") or {}
        readiness = manifest.get("readiness") or {}
        workbook = manifest.get("published_workbook") or {}
        revisions = manifest.get("input_revision") or {}
        chain = manifest.get("chain")

        record.update(
            {
                "release_id": release_id or None,
                "published_at": published_at or None,
                "commit_sha": commit.get("sha"),
                "gate_version": readiness.get("gate_version"),
                "published_workbook_sha256": workbook.get("sha256"),
                "published_workbook_stable_sha256": workbook.get(
                    "stable_sha256"
                ),
                "stock_source_etag": revisions.get("stock_source_etag"),
                "open_po_source_etag": revisions.get("open_po_source_etag"),
            }
        )

        published_dt = _parse_timestamp(published_at)
        record["_published_dt"] = published_dt
        if published_dt is None:
            _issue(
                record,
                "invalid_published_at",
                f"published_at không hợp lệ: {published_at!r}",
            )

        if not release_id:
            _issue(record, "missing_release_id", "release_id rỗng")
        elif path.name != f"{release_id}.json":
            _issue(
                record,
                "filename_release_id_mismatch",
                (
                    f"filename={path.name!r}; "
                    f"expected={release_id + '.json'!r}"
                ),
            )

        if chain is None:
            record["chain_status"] = "legacy_unlinked"
            _issue(
                record,
                "legacy_unlinked",
                "Manifest không có block chain; chỉ có thể audit nội tại.",
                severity="warning",
            )
        elif not isinstance(chain, dict):
            record["chain_status"] = "invalid"
            _issue(
                record,
                "invalid_chain",
                "chain phải là JSON object",
            )
        else:
            prev_id = str(
                chain.get("previous_release_id") or ""
            ).strip()
            prev_hash = str(
                chain.get("previous_manifest_sha256") or ""
            ).strip()
            record["previous_release_id"] = prev_id or None
            record["previous_manifest_sha256"] = prev_hash or None
            if bool(prev_id) != bool(prev_hash):
                record["chain_status"] = "invalid"
                _issue(
                    record,
                    "partial_chain_link",
                    (
                        "previous_release_id và previous_manifest_sha256 "
                        "phải cùng có hoặc cùng rỗng"
                    ),
                )
            elif prev_id:
                record["chain_status"] = "linked"
            else:
                record["chain_status"] = "anchor"

        try:
            verification = verify_nvl_release(
                path,
                strict=False,
                verify_commit=verify_commit,
                repo_root=repo_root,
            )
            record["verification_status"] = "passed"
            record["verification_warning_count"] = len(
                verification.get("warnings") or []
            )
        except NVLReleaseVerificationError as exc:
            record["verification_status"] = "failed"
            _issue(
                record,
                "manifest_verification_failed",
                str(exc),
            )

        records.append(record)

    valid_for_order = [
        record
        for record in records
        if record.get("_published_dt") is not None
        and record.get("release_id")
    ]
    valid_for_order.sort(
        key=lambda item: (
            item["_published_dt"],
            item["release_id"],
        )
    )

    by_id: dict[str, list[dict[str, Any]]] = {}
    for record in records:
        release_id = record.get("release_id")
        if release_id:
            by_id.setdefault(release_id, []).append(record)

    for release_id, same_id in by_id.items():
        if len(same_id) > 1:
            for record in same_id:
                _issue(
                    record,
                    "duplicate_release_id",
                    f"release_id {release_id!r} xuất hiện {len(same_id)} lần",
                )

    successor_map: dict[str, list[str]] = {}
    for index, record in enumerate(valid_for_order):
        prev_id = record.get("previous_release_id")
        prev_hash = record.get("previous_manifest_sha256")

        if record.get("chain_status") == "anchor":
            earlier_linked = [
                item
                for item in valid_for_order[:index]
                if item.get("chain_status") in {"anchor", "linked"}
            ]
            if earlier_linked:
                _issue(
                    record,
                    "chain_reset",
                    (
                        "Manifest anchor xuất hiện sau khi hash chain đã bắt đầu; "
                        "ledger bị reset."
                    ),
                )
                record["chain_status"] = "broken"
            continue

        if record.get("chain_status") != "linked":
            continue

        predecessor_candidates = by_id.get(prev_id or "", [])
        if not predecessor_candidates:
            _issue(
                record,
                "missing_predecessor",
                f"Không tìm thấy predecessor release {prev_id!r}",
            )
            record["chain_status"] = "broken"
            continue
        if len(predecessor_candidates) != 1:
            _issue(
                record,
                "ambiguous_predecessor",
                f"Predecessor {prev_id!r} không duy nhất",
            )
            record["chain_status"] = "broken"
            continue

        predecessor = predecessor_candidates[0]
        if predecessor.get("manifest_sha256") != prev_hash:
            _issue(
                record,
                "predecessor_hash_mismatch",
                (
                    f"expected={prev_hash}; "
                    f"actual={predecessor.get('manifest_sha256')}"
                ),
            )
            record["chain_status"] = "broken"

        previous_ordered = (
            valid_for_order[index - 1]
            if index > 0
            else None
        )
        if (
            previous_ordered is None
            or previous_ordered.get("release_id") != prev_id
        ):
            _issue(
                record,
                "chain_gap_or_branch",
                (
                    f"expected chronological predecessor="
                    f"{previous_ordered.get('release_id') if previous_ordered else None!r}; "
                    f"manifest points to={prev_id!r}"
                ),
            )
            record["chain_status"] = "broken"

        successor_map.setdefault(prev_id or "", []).append(
            record["release_id"]
        )

    for predecessor_id, successors in successor_map.items():
        if len(successors) > 1:
            for successor_id in successors:
                for record in by_id.get(successor_id, []):
                    _issue(
                        record,
                        "chain_fork",
                        (
                            f"Predecessor {predecessor_id!r} có nhiều successor: "
                            + ", ".join(sorted(successors))
                        ),
                    )
                    record["chain_status"] = "broken"

    latest_check: dict[str, Any] = {
        "path": str(latest_path),
        "status": "not_applicable" if not records else "unknown",
        "release_id": None,
        "detail": "",
    }

    if records:
        newest = valid_for_order[-1] if valid_for_order else None
        if not latest_path.is_file():
            latest_check.update(
                {
                    "status": "failed",
                    "detail": "Thiếu nvl/latest_release.json",
                }
            )
        else:
            latest_raw = latest_path.read_bytes()
            try:
                latest_manifest = _load_json_bytes(
                    latest_raw,
                    label="latest_release.json",
                )
                latest_id = str(
                    latest_manifest.get("release_id") or ""
                ).strip()
                latest_check["release_id"] = latest_id or None
                if newest is None:
                    latest_check.update(
                        {
                            "status": "failed",
                            "detail": (
                                "Không xác định được release mới nhất "
                                "từ ledger."
                            ),
                        }
                    )
                elif latest_id != newest.get("release_id"):
                    latest_check.update(
                        {
                            "status": "failed",
                            "detail": (
                                f"latest_release={latest_id!r}; "
                                f"newest={newest.get('release_id')!r}"
                            ),
                        }
                    )
                elif _sha256(latest_raw) != newest.get(
                    "manifest_sha256"
                ):
                    latest_check.update(
                        {
                            "status": "failed",
                            "detail": (
                                "latest_release.json không byte-identical "
                                "với immutable release file."
                            ),
                        }
                    )
                else:
                    latest_check.update(
                        {
                            "status": "passed",
                            "detail": "latest pointer khớp newest release",
                        }
                    )
            except ValueError as exc:
                latest_check.update(
                    {
                        "status": "failed",
                        "detail": str(exc),
                    }
                )

    for record in records:
        record["status"] = _record_status(record)

    error_count = sum(
        1
        for record in records
        for issue in record.get("issues", [])
        if issue.get("severity") == "error"
    )
    warning_count = sum(
        1
        for record in records
        for issue in record.get("issues", [])
        if issue.get("severity") == "warning"
    )
    if latest_check["status"] == "failed":
        error_count += 1

    ordered_records = sorted(
        records,
        key=lambda item: (
            item.get("_published_dt") or datetime.min.replace(
                tzinfo=timezone.utc
            ),
            item.get("release_id") or "",
            item.get("filename") or "",
        ),
    )
    for index, record in enumerate(ordered_records, start=1):
        record["sequence"] = index
        record.pop("_manifest", None)
        record.pop("_published_dt", None)

    overall = (
        "empty"
        if not records
        else ("failed" if error_count else "passed")
    )
    return {
        "schema": INDEX_SCHEMA,
        "schema_version": INDEX_SCHEMA_VERSION,
        "generated_at": datetime.now(timezone.utc).isoformat(),
        "status": overall,
        "releases_dir": str(releases_dir),
        "latest": latest_check,
        "summary": {
            "release_count": len(records),
            "passed_count": sum(
                1 for item in records if item["status"] == "passed"
            ),
            "warning_count": sum(
                1 for item in records if item["status"] == "warning"
            ),
            "failed_count": sum(
                1 for item in records if item["status"] == "failed"
            ),
            "issue_error_count": error_count,
            "issue_warning_count": warning_count,
            "legacy_unlinked_count": sum(
                1
                for item in records
                if item.get("chain_status") == "legacy_unlinked"
            ),
        },
        "releases": ordered_records,
    }


def write_index_json(
    index: dict[str, Any],
    path: str | Path,
) -> Path:
    path = Path(path)
    path.write_text(
        json.dumps(
            index,
            ensure_ascii=False,
            indent=2,
            sort_keys=True,
        )
        + "\n",
        encoding="utf-8",
    )
    return path


def write_index_csv(
    index: dict[str, Any],
    path: str | Path,
) -> Path:
    path = Path(path)
    fields = [
        "sequence",
        "published_at",
        "release_id",
        "commit_sha",
        "gate_version",
        "stock_source_etag",
        "open_po_source_etag",
        "published_workbook_sha256",
        "published_workbook_stable_sha256",
        "previous_release_id",
        "chain_status",
        "verification_status",
        "status",
        "issues",
    ]
    with path.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=fields)
        writer.writeheader()
        for record in index.get("releases", []):
            writer.writerow(
                {
                    field: (
                        " | ".join(
                            f"{issue.get('code')}: {issue.get('detail')}"
                            for issue in record.get("issues", [])
                        )
                        if field == "issues"
                        else record.get(field)
                    )
                    for field in fields
                }
            )
    return path
