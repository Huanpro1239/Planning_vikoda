"""Pure NVL code and quantity normalization rules."""

from __future__ import annotations

import math
import re
from typing import Any

EPS = 1e-6

def normalize_nvl_code(value: Any) -> str | None:
    """Chuẩn hóa mã vật tư NVL:

    - None, chuỗi rỗng hoặc boolean -> None
    - float với phần thập phân nguyên (ví dụ 12345.0) -> '12345'
    - int (ví dụ 12345) -> '12345'
    - str -> trim khoảng trắng đầu/cuối, giữ nguyên số 0 ở đầu (ví dụ '012345'),
      không chuyển đổi hoa/thường để bảo toàn đúng mã kế toán.
    """
    if value is None or isinstance(value, bool):
        return None

    if isinstance(value, (int, float)):
        if math.isnan(value) or math.isinf(value):
            return None
        if isinstance(value, float) and value.is_integer():
            return str(int(value))
        if isinstance(value, int):
            return str(value)
        return str(value).strip()

    text = str(value).strip()
    if not text:
        return None

    # Nếu chuỗi số float nguyên (ví dụ '12345.0') -> chuẩn hóa thành '12345'
    if re.match(r"^\d+\.0+$", text):
        return text.split(".")[0]

    return text


def parse_nvl_quantity(
    value: Any,
    convention: str = "strict",
    cell_name: str = "",
) -> float:
    """Phân tích và kiểm tra giá trị số lượng tồn kho NVL.

    Quy tắc:
    - Giá trị số (int, float) nguyên bản từ Excel được chép trực tiếp.
    - Cho phép số 0, số âm, số thập phân lẻ.
    - None, chuỗi rỗng, boolean, lỗi Excel, NaN, Infinity -> raise ValueError.
    - Chuỗi văn bản phải tuân thủ phân nhóm số nghiêm ngặt theo quy ước:
        + 'vi': Dấu chấm phân nhóm hàng nghìn (1.234.567), dấu phẩy thập phân (1,5; 1.234,56).
        + 'en': Dấu phẩy phân nhóm hàng nghìn (1,234,567), dấu chấm thập phân (1.5; 1,234.56).
        + 'strict': Nhận diện định dạng không mơ hồ (cả hai dấu phân cách hoặc số chữ số lẻ != 3).
                    Từ chối các chuỗi mơ hồ ('1,234' hoặc '1.234') và từ chối phân nhóm sai ('1,2,3').
    """
    prefix = f"[{cell_name}] " if cell_name else ""

    if value is None or (isinstance(value, str) and value.strip() == ""):
        raise ValueError(f"{prefix}Ô số lượng tồn kho rỗng.")

    if isinstance(value, bool):
        raise ValueError(f"{prefix}Chứa giá trị boolean ({value}), không phải số.")

    # 1. Số dạng numeric trong Excel được nhận trực tiếp
    if isinstance(value, (int, float)):
        fval = float(value)
        if math.isnan(fval) or math.isinf(fval):
            raise ValueError(f"{prefix}Chứa giá trị NaN hoặc vô cực.")
        return fval

    # 2. Xử lý chuỗi văn bản
    if not isinstance(value, str):
        raise ValueError(f"{prefix}Kiểu dữ liệu không được hỗ trợ: {type(value)}")

    cleaned = value.strip()
    if not cleaned:
        raise ValueError(f"{prefix}Chuỗi rỗng không phải số hợp lệ.")

    # Kiểm tra dấu âm/dương
    sign = 1.0
    if cleaned.startswith("-"):
        sign = -1.0
        cleaned = cleaned[1:].strip()
    elif cleaned.startswith("+"):
        cleaned = cleaned[1:].strip()

    if not cleaned:
        raise ValueError(f"{prefix}Chuỗi số không hợp lệ: {value!r}")

    # Chuỗi số nguyên thuần túy
    if cleaned.isdigit():
        return sign * float(cleaned)

    # Đếm dấu phân cách
    dot_count = cleaned.count(".")
    comma_count = cleaned.count(",")

    if dot_count == 0 and comma_count == 0:
        raise ValueError(f"{prefix}Giá trị {value!r} không thể chuyển đổi thành số tồn hợp lệ.")

    valid_chars = set("0123456789.,")
    if not set(cleaned).issubset(valid_chars):
        raise ValueError(f"{prefix}Chuỗi số chứa ký tự không hợp lệ: {value!r}")

    # Chặn phân cách nhóm liên tiếp: '..', ',,', '.,', ',.'
    if re.search(r"[.,]{2,}", cleaned):
        raise ValueError(f"{prefix}Phân cách nhóm số không hợp lệ: {value!r}")

    conv = (convention or "strict").strip().lower()

    # Trường hợp 1: Chứa cả dấu chấm và dấu phẩy
    if dot_count > 0 and comma_count > 0:
        is_en_pattern = bool(re.match(r"^\d{1,3}(,\d{3})+(\.\d+)$", cleaned))
        is_vi_pattern = bool(re.match(r"^\d{1,3}(\.\d{3})+(,\d+)$", cleaned))

        if is_en_pattern:
            if conv == "vi":
                raise ValueError(f"{prefix}Định dạng số kiểu Anh {value!r} không khớp với quy ước Việt Nam đã cấu hình.")
            return sign * float(cleaned.replace(",", ""))
        elif is_vi_pattern:
            if conv == "en":
                raise ValueError(f"{prefix}Định dạng số kiểu Việt Nam {value!r} không khớp với quy ước Anh đã cấu hình.")
            return sign * float(cleaned.replace(".", "").replace(",", "."))
        else:
            raise ValueError(f"{prefix}Phân nhóm dấu phân cách số không đúng quy cách: {value!r}")

    # Trường hợp 2: Chỉ chứa dấu phẩy
    if comma_count > 0:
        if conv == "en":
            # Trong chuẩn EN, dấu phẩy chỉ có thể là phân nhóm hàng nghìn (mỗi nhóm đúng 3 chữ số)
            if not re.match(r"^\d{1,3}(,\d{3})+$", cleaned):
                raise ValueError(f"{prefix}Phân nhóm dấu phẩy không đúng quy cách số tiếng Anh: {value!r}")
            return sign * float(cleaned.replace(",", ""))
        elif conv == "vi":
            if comma_count > 1:
                raise ValueError(f"{prefix}Dấu phẩy phân nhóm không hợp lệ trong quy ước Việt Nam: {value!r}")
            if not re.match(r"^\d+,\d+$", cleaned):
                raise ValueError(f"{prefix}Dấu phẩy thập phân không đúng định dạng: {value!r}")
            return sign * float(cleaned.replace(",", "."))
        else:  # strict / auto
            if comma_count > 1:
                raise ValueError(f"{prefix}Nhiều dấu phẩy không hợp lệ trong chế độ nghiêm ngặt: {value!r}")
            m = re.match(r"^(\d+),(\d+)$", cleaned)
            if not m:
                raise ValueError(f"{prefix}Định dạng số chứa dấu phẩy không hợp lệ: {value!r}")
            if len(m.group(2)) == 3:
                raise ValueError(
                    f"{prefix}Chuỗi số mơ hồ giữa phân nhóm hàng nghìn và số thập phân: {value!r}. "
                    "Vui lòng cấu hình number_convention ('vi' hoặc 'en')."
                )
            # Số chữ số sau dấu phẩy != 3 -> chắc chắn là dấu phẩy thập phân
            return sign * float(cleaned.replace(",", "."))

    # Trường hợp 3: Chỉ chứa dấu chấm
    if dot_count > 0:
        if conv == "vi":
            # Trong chuẩn VI, dấu chấm chỉ có thể là phân nhóm hàng nghìn
            if not re.match(r"^\d{1,3}(\.\d{3})+$", cleaned):
                raise ValueError(f"{prefix}Phân nhóm dấu chấm không đúng quy cách Việt Nam: {value!r}")
            return sign * float(cleaned.replace(".", ""))
        elif conv == "en":
            if dot_count > 1:
                raise ValueError(f"{prefix}Nhiều dấu chấm không hợp lệ trong quy ước tiếng Anh: {value!r}")
            if not re.match(r"^\d+\.\d+$", cleaned):
                raise ValueError(f"{prefix}Dấu chấm thập phân không đúng định dạng: {value!r}")
            return sign * float(cleaned)
        else:  # strict / auto
            if dot_count > 1:
                raise ValueError(f"{prefix}Nhiều dấu chấm không hợp lệ trong chế độ nghiêm ngặt: {value!r}")
            m = re.match(r"^(\d+)\.(\d+)$", cleaned)
            if not m:
                raise ValueError(f"{prefix}Định dạng số chứa dấu chấm không hợp lệ: {value!r}")
            if len(m.group(2)) == 3:
                raise ValueError(
                    f"{prefix}Chuỗi số mơ hồ giữa phân nhóm hàng nghìn và số thập phân: {value!r}. "
                    "Vui lòng cấu hình number_convention ('vi' hoặc 'en')."
                )
            # Số chữ số sau dấu chấm != 3 -> chắc chắn là dấu chấm thập phân
            return sign * float(cleaned)

    raise ValueError(f"{prefix}Không thể xử lý giá trị số: {value!r}")


def _is_finite_number(val: Any) -> bool:
    """Kiểm tra một giá trị có phải là số hữu hạn (không phải bool, nan, inf)."""
    if val is None or isinstance(val, bool):
        return False
    if isinstance(val, (int, float)):
        return not (math.isnan(val) or math.isinf(val))
    if isinstance(val, str):
        try:
            f = float(val)
            return not (math.isnan(f) or math.isinf(f))
        except ValueError:
            return False
    return False


def _cells_match(actual: Any, expected: Any, eps: float = EPS) -> bool:
    """So sánh hai giá trị ô mà không bắt buộc ô phi số phải chuyển thành float."""
    # Cả hai là None hoặc chuỗi rỗng
    is_act_empty = actual is None or (isinstance(actual, str) and actual.strip() == "")
    is_exp_empty = expected is None or (isinstance(expected, str) and expected.strip() == "")
    if is_act_empty and is_exp_empty:
        return True
    if is_act_empty != is_exp_empty:
        return False

    # Cả hai là số hữu hạn: so sánh có dung sai EPS
    if _is_finite_number(actual) and _is_finite_number(expected):
        return abs(float(actual) - float(expected)) <= eps

    # So sánh chuỗi hoặc giá trị trực tiếp (ví dụ '-', 'N/A', ghi chú text)
    return str(actual).strip() == str(expected).strip()


__all__ = ["EPS", "normalize_nvl_code", "parse_nvl_quantity"]
