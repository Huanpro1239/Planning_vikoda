"""Chạy toàn bộ engine kế hoạch ở chế độ offline, không cần SharePoint.

Dùng khi muốn kiểm tra/nghiệm thu proposal từ các bản sao tải về máy:

    python run_offline.py \
        --target "Sắp kế hoạch.xlsx" \
        --actual "Bao cao ton thuc te hien tai.xlsx" \
        --factory-vikoda NXT_Vikoda.xlsm \
        --factory-vkd NXT_VKD.xlsm \
        --accounting-vikoda XNT_ketoan_Vikoda.xlsm \
        --accounting-vkd XNT_ketoan_VKD.xlsm \
        --out ./out --verify

Script chỉ ĐỌC các file nguồn và GHI proposal + report vào thư mục ``--out``.
Không kết nối mạng, không upload, không đổi runtime state gốc (trừ khi
``--save-runtime-state`` được bật). Đây là con đường an toàn để tái lập đúng số
liệu mà GitHub Actions sẽ tạo ra.
"""
from __future__ import annotations

import argparse
import hashlib
import json
from pathlib import Path

import sync_planning_metrics as metrics
from planning_pipeline import SOURCE_KEYS, prepare_pipeline_output
from planning_schedule_report import print_operational_report, save_schedule_report


_ARG_TO_KEY = {
    "actual": "actual_stock",
    "factory_vikoda": "factory_vikoda",
    "factory_vkd": "factory_vkd",
    "accounting_vikoda": "accounting_vikoda",
    "accounting_vkd": "accounting_vkd",
}


def _read_bytes(path: str, label: str) -> bytes:
    file_path = Path(path)
    if not file_path.is_file():
        raise SystemExit(f"[OFFLINE] Không tìm thấy file {label}: {path}")
    return file_path.read_bytes()


def _revision(target_bytes: bytes, source_bytes: dict[str, bytes]) -> dict:
    """Provenance tối thiểu để verify chấp nhận (không cần ETag SharePoint)."""
    return {
        "mode": "offline_local_files",
        "target": {"sha256": hashlib.sha256(target_bytes).hexdigest()},
        "sources": {
            key: {"sha256": hashlib.sha256(source_bytes[key]).hexdigest()}
            for key in SOURCE_KEYS
        },
    }


def _parse_args(argv=None):
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--target", required=True, help="Sắp kế hoạch.xlsx")
    parser.add_argument("--actual", required=True, help="Bao cao ton thuc te hien tai.xlsx")
    parser.add_argument("--factory-vikoda", required=True, dest="factory_vikoda")
    parser.add_argument("--factory-vkd", required=True, dest="factory_vkd")
    parser.add_argument("--accounting-vikoda", required=True, dest="accounting_vikoda")
    parser.add_argument("--accounting-vkd", required=True, dest="accounting_vkd")
    parser.add_argument("--out", default="./offline_out", help="Thư mục xuất proposal/report")
    parser.add_argument(
        "--runtime-state",
        default=None,
        help="Đường dẫn planning_runtime.json thay thế (mặc định lấy file trong repo).",
    )
    parser.add_argument(
        "--save-runtime-state",
        action="store_true",
        help="Ghi lại runtime state đề xuất (mặc định KHÔNG ghi để giữ nguyên gốc).",
    )
    parser.add_argument(
        "--verify",
        action="store_true",
        help="Chạy verify_planning_month trên proposal vừa tạo.",
    )
    return parser.parse_args(argv)


def main(argv=None):
    args = _parse_args(argv)

    target_bytes = _read_bytes(args.target, "target")
    source_bytes = {
        key: _read_bytes(getattr(args, arg), key)
        for arg, key in _ARG_TO_KEY.items()
    }

    if args.runtime_state:
        runtime_state = json.loads(Path(args.runtime_state).read_text(encoding="utf-8"))
    else:
        runtime_state = metrics.load_runtime_state()

    final_bytes, report, next_state = prepare_pipeline_output(
        target_bytes,
        source_bytes,
        runtime_state=runtime_state,
        input_revision=_revision(target_bytes, source_bytes),
    )

    out_dir = Path(args.out)
    out_dir.mkdir(parents=True, exist_ok=True)
    proposal_path = out_dir / "planning_proposal.xlsx"
    proposal_path.write_bytes(final_bytes)
    save_schedule_report(report, out_dir / "planning_schedule_report.json")

    print_operational_report(report)
    print(f"[OFFLINE] Proposal: {proposal_path}")
    print(f"[OFFLINE] Report:   {out_dir / 'planning_schedule_report.json'}")

    if args.save_runtime_state:
        state_path = out_dir / "planning_runtime.json"
        state_path.write_text(
            json.dumps(next_state, ensure_ascii=False, indent=2) + "\n",
            encoding="utf-8",
        )
        print(f"[OFFLINE] Runtime state đề xuất: {state_path}")

    if args.verify:
        # Import trễ để chế độ không verify không phải nạp thêm.
        from verify_planning_month import verify_workbook

        verify_workbook(final_bytes, schedule_report=report)
        print("[OFFLINE] verify_planning_month: PASS")

    return {"report": report, "proposal_path": str(proposal_path)}


if __name__ == "__main__":
    main()
