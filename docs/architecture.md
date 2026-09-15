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
