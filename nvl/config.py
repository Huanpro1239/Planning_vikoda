"""NVL configuration loading."""

from __future__ import annotations

import json
from pathlib import Path

from nvl.models import NVLConfig

DEFAULT_CONFIG_FILE = "nvl_stock_config.json"

def load_nvl_config(config_path: str | Path | None = None) -> NVLConfig:
    """Đọc cấu hình đồng bộ NVL từ file JSON."""
    if config_path is None:
        config_path = Path(__file__).resolve().parents[1] / DEFAULT_CONFIG_FILE
    else:
        config_path = Path(config_path)

    if not config_path.exists():
        raise FileNotFoundError(f"Không tìm thấy file cấu hình: {config_path}")

    with open(config_path, "r", encoding="utf-8") as f:
        data = json.load(f)

    src = data.get("source", {})
    tgt = data.get("target", {})
    pol = data.get("policies", {})

    return NVLConfig(
        source_name=src.get("name", "XNT_ketoan_Vikoda.xlsm"),
        source_path=src.get("sharepoint_path", ""),
        source_sheet=src.get("sheet_name", "Sheet1"),
        source_code_col=int(src.get("code_column", 2)),
        source_code_col_letter=src.get("code_column_letter", "B"),
        source_value_col=int(src.get("value_column", 13)),
        source_value_col_letter=src.get("value_column_letter", "M"),
        source_start_row=int(src.get("start_row", 2)),
        source_sourcedoc=src.get("sourcedoc", ""),
        target_name=tgt.get("name", "Kế hoạch mua hàng.xlsx"),
        target_path=tgt.get("sharepoint_path", ""),
        target_sheet=tgt.get("sheet_name", "Ton_NVL"),
        target_code_col=int(tgt.get("code_column", 1)),
        target_code_col_letter=tgt.get("code_column_letter", "A"),
        target_value_col=int(tgt.get("value_column", 4)),
        target_value_col_letter=tgt.get("value_column_letter", "D"),
        target_start_row=int(tgt.get("start_row", 2)),
        target_sourcedoc=tgt.get("sourcedoc", ""),
        direct_copy=bool(pol.get("direct_copy", True)),
        allow_zero=bool(pol.get("allow_zero", True)),
        allow_negative=bool(pol.get("allow_negative", True)),
        preserve_missing_in_source=bool(pol.get("preserve_missing_in_source", True)),
        reject_duplicates=bool(pol.get("reject_duplicates", True)),
        reject_formula_in_target_cell=bool(pol.get("reject_formula_in_target_cell", True)),
        number_convention=str(pol.get("number_convention", "strict")).strip().lower(),
    )


__all__ = ["DEFAULT_CONFIG_FILE", "load_nvl_config"]
