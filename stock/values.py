"""Pure value normalization for finished-goods stock."""

import re


CODE_PATTERN = re.compile(r"^\d{6,}$")


def normalize_code(value, vkd_to_vikoda=False):
    if value is None:
        return None

    if isinstance(value, float) and value.is_integer():
        value = int(value)

    text = str(value).strip()
    if not CODE_PATTERN.fullmatch(text):
        return None

    if vkd_to_vikoda and text.startswith("2"):
        text = "1" + text[1:]

    return text


def to_number(value, cell_name="", default=0):
    if value is None or value == "":
        return default

    if isinstance(value, bool):
        raise ValueError(
            f"{cell_name} chứa TRUE/FALSE, không phải số."
        )

    if isinstance(value, (int, float)):
        return value

    text = str(value).strip().replace(",", "")
    try:
        return float(text)
    except ValueError as exc:
        raise ValueError(
            f"{cell_name} có giá trị không phải số: {value!r}"
        ) from exc


def clean_number(value):
    if isinstance(value, float) and value.is_integer():
        return int(value)
    return value
