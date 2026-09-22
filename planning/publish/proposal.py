"""Stable proposal identity and review artifacts."""

import hashlib
import json
from io import BytesIO
from zipfile import BadZipFile, ZipFile

from planning_schedule_report import save_schedule_report

from .constants import (
    AUDIT_REVISION_FILE,
    PROPOSAL_WORKBOOK_FILE,
    VOLATILE_XLSX_PARTS,
)


def proposal_output_sha256(final_bytes):
    """Hash stable XLSX payload while ignoring volatile Office metadata."""
    raw_sha256 = hashlib.sha256(final_bytes).hexdigest()
    try:
        with ZipFile(BytesIO(final_bytes), "r") as archive:
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
                data = archive.read(name)
                digest.update(len(name_bytes).to_bytes(4, "big"))
                digest.update(name_bytes)
                digest.update(len(data).to_bytes(8, "big"))
                digest.update(data)
            return digest.hexdigest()
    except (BadZipFile, OSError):
        return raw_sha256


def proposal_id(report, final_bytes):
    output_sha256 = hashlib.sha256(final_bytes).hexdigest()
    stable_output_sha256 = proposal_output_sha256(final_bytes)
    payload = {
        "algorithm": report.get("algorithm"),
        "plan_month": report.get("plan_month"),
        "input_revision": report.get("input_revision") or {},
        "output_sha256": stable_output_sha256,
    }
    canonical = json.dumps(
        payload,
        ensure_ascii=False,
        sort_keys=True,
        separators=(",", ":"),
    ).encode("utf-8")
    return (
        hashlib.sha256(canonical).hexdigest(),
        output_sha256,
        stable_output_sha256,
    )


def with_proposal_identity(report, final_bytes):
    result = dict(report)
    identity, output_sha256, stable_output_sha256 = proposal_id(
        result,
        final_bytes,
    )
    result["output_sha256"] = output_sha256
    result["proposal_output_sha256"] = stable_output_sha256
    result["proposal_id"] = identity
    return result


def save_proposal_artifacts(final_bytes, report):
    PROPOSAL_WORKBOOK_FILE.write_bytes(final_bytes)
    save_schedule_report(report)
    AUDIT_REVISION_FILE.write_text(
        json.dumps(
            report.get("input_revision") or {},
            ensure_ascii=False,
            indent=2,
        )
        + "\n",
        encoding="utf-8",
    )
