# Kiến trúc mục tiêu

Repo đang được refactor theo hướng tách nghiệp vụ khỏi hạ tầng nhưng vẫn giữ backward compatibility cho workflow production.

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
2. Entry-point cũ ở root được giữ trong giai đoạn chuyển tiếp để GitHub Actions và công cụ nội bộ không gãy.
3. Code mới ưu tiên import qua package canonical (`sharepoint.client`, `nvl.*`, `planning.*`).
4. Mỗi bước migrate phải qua toàn bộ test suite trước khi merge.
5. Sau khi toàn bộ caller đã chuyển sang package mới, root module mới được thu gọn thành compatibility shim hoặc xóa ở một major cleanup riêng.

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
