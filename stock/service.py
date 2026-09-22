"""Standalone SharePoint orchestration for finished-goods stock."""

from sharepoint.client import GraphClient, get_access_token

from .constants import (
    DEST_PATH,
    SOURCE_ACCOUNTING_VIKODA_PATH,
    SOURCE_ACCOUNTING_VKD_PATH,
    SOURCE_ACTUAL_PATH,
    SOURCE_FACTORY_VIKODA_PATH,
    SOURCE_FACTORY_VKD_PATH,
)
from .readers import (
    read_actual_stock,
    read_conversion_factors,
    read_single_value_source,
)
from .state import SYNC_VERSION, load_state, save_state
from .workbook import patch_destination_workbook


def main():
    token = get_access_token()
    graph = GraphClient(token)

    site_id = graph.get_site_id()
    drive_id = graph.get_default_drive_id(site_id)

    print("Đã kết nối SharePoint Planning.")

    sources = {
        "actual_stock": {
            "path": SOURCE_ACTUAL_PATH,
            "label": "Tồn thực tế",
        },
        "factory_vikoda": {
            "path": SOURCE_FACTORY_VIKODA_PATH,
            "label": "Tồn nhà máy Vikoda",
        },
        "factory_vkd": {
            "path": SOURCE_FACTORY_VKD_PATH,
            "label": "Tồn nhà máy VKD",
        },
        "accounting_vikoda": {
            "path": SOURCE_ACCOUNTING_VIKODA_PATH,
            "label": "Tồn kế toán Vikoda",
        },
        "accounting_vkd": {
            "path": SOURCE_ACCOUNTING_VKD_PATH,
            "label": "Tồn kế toán VKD",
        },
    }

    for source in sources.values():
        source["item"] = graph.get_item_by_path(
            drive_id,
            source["path"],
        )

    current_etags = {
        key: source["item"]["eTag"]
        for key, source in sources.items()
    }

    dest_item = graph.get_item_by_path(
        drive_id,
        DEST_PATH,
    )
    dest_bytes = graph.download_file(
        drive_id,
        dest_item["id"],
    )

    conversion_factors, conversion_hash = read_conversion_factors(
        dest_bytes
    )

    old_state = load_state()
    old_etags = old_state.get("sources", {})

    changed_sources = [
        key
        for key, etag in current_etags.items()
        if old_etags.get(key) != etag
    ]

    if old_state.get("sync_version") != SYNC_VERSION:
        changed_sources.append("logic_version")

    if old_state.get("conversion_hash") != conversion_hash:
        changed_sources.append("Danh_muc!I")

    for source in sources.values():
        print(
            f"[{source['label']}] sửa lần cuối:",
            source["item"].get("lastModifiedDateTime"),
        )

    if not changed_sources:
        print(
            "Không có file nguồn/quy cách nào thay đổi. Kết thúc."
        )
        return

    print(
        "Nguồn hoặc logic thay đổi: "
        + ", ".join(changed_sources)
    )

    source_bytes = {
        key: graph.download_file(
            drive_id,
            source["item"]["id"],
        )
        for key, source in sources.items()
    }

    actual_stock = read_actual_stock(
        source_bytes["actual_stock"]
    )
    factory_vikoda = read_single_value_source(
        source_bytes["factory_vikoda"],
        label="Tồn nhà máy Vikoda",
        source_name="NXT_Vikoda.xlsm",
        sheet_name="Sheet1",
        code_column=2,
        value_column=12,
        value_column_letter="L",
    )
    factory_vkd = read_single_value_source(
        source_bytes["factory_vkd"],
        label="Tồn nhà máy VKD",
        source_name="NXT_VKD.xlsm",
        sheet_name="Sheet1",
        code_column=2,
        value_column=12,
        value_column_letter="L",
        vkd_to_vikoda=True,
    )
    accounting_vikoda = read_single_value_source(
        source_bytes["accounting_vikoda"],
        label="Tồn kế toán Vikoda",
        source_name="XNT_ketoan_Vikoda.xlsm",
        sheet_name="Sheet1",
        code_column=2,
        value_column=13,
        value_column_letter="M",
    )
    accounting_vkd = read_single_value_source(
        source_bytes["accounting_vkd"],
        label="Tồn kế toán VKD",
        source_name="XNT_ketoan_VKD.xlsm",
        sheet_name="Sheet1",
        code_column=2,
        value_column=13,
        value_column_letter="M",
        vkd_to_vikoda=True,
    )

    updated_dest_bytes = patch_destination_workbook(
        dest_bytes,
        actual_stock=actual_stock,
        factory_vikoda=factory_vikoda,
        factory_vkd=factory_vkd,
        accounting_vikoda=accounting_vikoda,
        accounting_vkd=accounting_vkd,
        conversion_factors=conversion_factors,
    )

    result = graph.upload_file(
        drive_id,
        dest_item["id"],
        updated_dest_bytes,
        expected_etag=dest_item["eTag"],
    )

    print(
        "Upload thành công:",
        result.get("name", "Sắp kế hoạch.xlsx"),
    )

    save_state(current_etags, conversion_hash)
    print("SYNC THÀNH CÔNG.")
