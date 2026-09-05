import hashlib
import json
from io import BytesIO

from openpyxl import load_workbook

import sync_stock


def read_conversion_factors_robust(dest_bytes):
    workbook = load_workbook(
        BytesIO(dest_bytes),
        data_only=True,
        read_only=True,
    )

    try:
        if sync_stock.MASTER_SHEET not in workbook.sheetnames:
            raise RuntimeError(
                f"Không tìm thấy sheet {sync_stock.MASTER_SHEET!r} trong file đích."
            )

        worksheet = workbook[sync_stock.MASTER_SHEET]
        factors = {}

        for row_number, values in enumerate(
            worksheet.iter_rows(min_row=2, max_col=9, values_only=True),
            start=2,
        ):
            code = sync_stock.normalize_code(values[0] if values else None)
            if not code:
                continue

            if code in factors:
                raise RuntimeError(
                    f"[{sync_stock.MASTER_SHEET}] Mã {code} bị lặp trong cột A."
                )

            factor = sync_stock.to_number(
                values[8] if len(values) >= 9 else None,
                f"{sync_stock.MASTER_SHEET}!I{row_number}",
            )
            if factor <= 0:
                raise RuntimeError(
                    f"[{sync_stock.MASTER_SHEET}] Quy cách của mã {code} "
                    f"phải > 0, hiện là {factor!r}."
                )

            factors[code] = factor

        if not factors:
            raise RuntimeError(
                f"[{sync_stock.MASTER_SHEET}] Không đọc được mã/quy cách từ A:I."
            )

        payload = json.dumps(
            {k: factors[k] for k in sorted(factors)},
            ensure_ascii=False,
            separators=(",", ":"),
        ).encode("utf-8")
        conversion_hash = hashlib.sha256(payload).hexdigest()

        print(
            f"[{sync_stock.MASTER_SHEET}] Đọc {len(factors)} quy cách "
            "bằng chế độ tương thích sheet không có dimension."
        )
        return factors, conversion_hash
    finally:
        workbook.close()


sync_stock.read_conversion_factors = read_conversion_factors_robust


if __name__ == "__main__":
    sync_stock.main()
