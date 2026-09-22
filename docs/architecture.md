# Kiến trúc mục tiêu

Repo đã hoàn tất các bước refactor chính theo hướng tách nghiệp vụ khỏi hạ tầng; workflow production dùng trực tiếp các entrypoint hiện hành.

```text
planning/     # planning engine, scheduling, proposal/publish orchestration
nvl/          # tồn NVL, PO mở, purchase planning
sharepoint/   # Microsoft Graph / SharePoint client, identity, retry, ETag
excel/        # workbook/XML read-write-verify utilities
domain/       # business rules thuần, không I/O
scripts/      # công cụ vận hành/audit
tests/        # regression/integration tests
docs/         # architecture, review history, runbooks
```

## Quy tắc migrate

1. Không đổi nghiệp vụ trong PR cấu trúc.
2. Compatibility shim chỉ được giữ khi còn consumer thực tế; shim không còn consumer phải được xóa.
3. Code mới ưu tiên import qua package canonical (`sharepoint.client`, `nvl.*`, `planning.*`).
4. Mỗi bước migrate phải qua toàn bộ test suite trước khi merge.
5. Sau khi toàn bộ caller đã chuyển sang package mới, compatibility shim phải được xóa và CI chặn tái xuất hiện.

## Thứ tự refactor

- Phase A: package layout + di chuyển docs.
- Phase B: hợp nhất SharePoint/Graph client.
- Phase C: migrate NVL modules.
- Phase D: migrate planning modules.
- Phase E: xóa compatibility shim sau khi workflow/tests không còn phụ thuộc.


## Repository hygiene

- `scripts/` chỉ chứa công cụ vận hành/audit dùng cho production hoặc support.
- `scripts/dev/` chứa công cụ chẩn đoán thủ công dành cho developer.
- `docs/archive/` chứa kế hoạch/review lịch sử; không để tài liệu tạm ở repository root.
- Package canonical không dùng `import *`; chỉ export API ổn định qua `__all__`.
- CI compile toàn bộ repository trước khi chạy unit tests.


## NVL module boundaries

`sync_nvl_stock.py` is a compatibility CLI only. Canonical responsibilities:

- `nvl/models.py`: immutable data contracts/state containers.
- `nvl/config.py`: JSON configuration loading.
- `nvl/values.py`: pure code/quantity normalization.
- `nvl/reconcile.py`: source reading and target reconciliation.
- `nvl/workbook.py`: OpenXML patching and integrity verification.
- `nvl/reporting.py`: audit/error report serialization.
- `nvl/service.py`: SharePoint orchestration, retries and publish boundary.
- `excel/workbook_xml.py`: reusable low-level OpenXML helpers.
- `excel/openpyxl_io.py`: deterministic workbook/ZipFile lifecycle cleanup.

NVL runtime modules must not import `sync_stock.py` directly. SharePoint access goes
through `sharepoint.client`, while workbook XML utilities go through `excel/`.


## Canonical SharePoint dependency

`sharepoint/client.py` owns `GraphClient`, `GraphRequestError`, Graph constants,
retry classification and SharePoint identity/read helpers. Authentication comes from
`sharepoint/auth.py`.

Legacy `sync_stock.py` may re-export the canonical Graph API for compatibility, but
canonical/runtime callers must import Graph/auth symbols from `sharepoint.client`.
The dependency direction is therefore:

```text
planning/ nvl/ scripts/ -> sharepoint.client -> sharepoint.auth
sync_stock.py -----------^  (compatibility re-export only)
```

Production workflows call their entrypoints directly. `scripts/auth_runner.py` has been removed; authentication is resolved by `sharepoint.auth` through canonical imports.


## Canonical Planning dependency

Planning implementations now live under the `planning/` package:

- `planning/metrics/`: static metrics package:
  - `state.py`: runtime-state serialization and period helpers.
  - `readers.py`: actual stock, system receipts, No kho and Danh_muc readers.
  - `calculation.py`: explicit all-month M:R calculations.
  - `workbook.py`: M:R OpenXML patch/diff.
  - `service.py`: standalone SharePoint orchestration.
- `planning/khsx_ki.py`: KHSX_ki weekly aggregation and verification.
- `planning/pipeline.py`: pure proposal-building pipeline.

Legacy metrics patchers `sync_planning_metrics_compat.py`, `sync_planning_metrics_direct.py` and `sync_planning_metrics_all_months.py` are removed. No module may rebind `planning.metrics` functions or policy state at import/runtime. No kho debt, Danh_muc leadtime and all-month calculations are wired explicitly through function arguments.
