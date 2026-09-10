# Kế hoạch bổ sung đồng bộ tồn nguyên vật liệu — giao Antigravity

Ngày phân tích: 09/09/2026. Repository: https://github.com/Huanpro1239/Planning_vikoda

Đã fetch và đối chiếu `origin/main` tại commit `3086495b51d34c94c3959f39b3ddbe1630228918`. Các module tồn kho, Graph, pipeline và workflow nêu dưới đây giống bản checkout đã đọc. Chỉ lập kế hoạch; chưa thay đổi code chạy, chưa ghi SharePoint hoặc push GitHub.

## 1. Yêu cầu đã chốt

| Nội dung | Nguồn | Đích |
|---|---|---|
| Workbook | `XNT_ketoan_Vikoda.xlsm` | `Kế hoạch mua hàng.xlsx` |
| Sheet | `Sheet1` | `Ton_NVL` |
| Cột mã | B | A |
| Cột số tồn | M | D — người dùng đã xác nhận |
| SharePoint sourcedoc | `C0372C81-A402-4A11-B9E0-847768CC9CFB` | `89D1BA7B-006E-4527-B879-ABF120309214` |

Với mỗi dòng vật tư hợp lệ ở đích, tìm dòng nguồn có cùng mã và gán `Ton_NVL!D = Sheet1!M`. Ghép theo mã, không theo thứ tự dòng. Chép nguyên số lượng, không chia quy cách, không trừ tồn nhà máy, không đổi tiền tố mã 2 thành 1.

Liên kết nguồn: https://vikodacomvn.sharepoint.com/:x:/r/sites/Planning/_layouts/15/Doc.aspx?sourcedoc=%7BC0372C81-A402-4A11-B9E0-847768CC9CFB%7D&file=XNT_ketoan_Vikoda.xlsm&action=default&mobileredirect=true

Liên kết đích: https://vikodacomvn.sharepoint.com/:x:/r/sites/Planning/_layouts/15/Doc.aspx?sourcedoc=%7B89D1BA7B-006E-4527-B879-ABF120309214%7D&file=K%E1%BA%BF%20ho%E1%BA%A1ch%20mua%20h%C3%A0ng.xlsx&action=default&mobileredirect=true

Chưa đọc được nội dung hai workbook qua liên kết trong phiên phân tích. Trước chạy thật cần xác minh drive/path hoặc driveItem ID của file đích, dòng bắt đầu dữ liệu, tiêu đề D, dạng mã, đơn vị tính và tình trạng công thức M/D. Không đoán thư mục đích từ tên file; GUID `sourcedoc` không được gán trực tiếp làm Graph driveItem ID.

## 2. Kết quả phân tích code hiện có

- `sync_stock.py`: có `GraphClient`, `get_access_token()`, tải file, upload có `If-Match`, phân loại lỗi retry, các helper XML/ZIP. `SOURCE_ACCOUNTING_VIKODA_PATH` đã trỏ đến `Tinh san xuat Mua hang 2027/Ton He thong/Ton Ke Toan/XNT_ketoan_Vikoda.xlsm`.
- `read_single_value_source()` và `planning_pipeline.py` đã đọc `Sheet1`, mã B, số lượng M của chính nguồn này.
- `patch_destination_workbook()` hiện chỉ phục vụ `Sắp kế hoạch.xlsx/Ton_kho`: tính G bằng tồn kế toán chia quy cách rồi trừ tồn nhà máy. Không dùng hàm này cho NVL.
- `normalize_code()` hiện chỉ nhận `^\d{6,}$`; `to_number()` coi ô trống là 0 và bỏ dấu phẩy. Không áp dụng nguyên trạng cho NVL khi chưa biết dạng mã và số liệu.
- `.github/workflows/sync-stock.yml` chạy `sync_planning_pipeline.py`, có lịch 06:00 Việt Nam và sự kiện `sharepoint_stock_updated`; đây là pipeline sản xuất với proposal/publish riêng.
- `graph_retry.py` hỗ trợ `Retry-After`. Có thể tái sử dụng cùng cơ chế tải lại snapshot sau xung đột.

## 3. Kiến trúc đề xuất

Thêm tiến trình NVL độc lập trong cùng repository, tái sử dụng helper kết nối và XML thuần. Không đưa workbook mua hàng vào pipeline tính kế hoạch sản xuất: hai file đích cần được đồng bộ độc lập, tránh NVL bị chặn bởi thiếu công suất sản xuất và tránh xuất hiện giao dịch nửa chừng giữa hai workbook.

| File | Công việc |
|---|---|
| `sync_nvl_stock.py` | Reader NVL, mapping, patch `Ton_NVL!D`, báo cáo, CLI offline/online, retry và publish |
| `nvl_stock_config.json` | Cấu hình nguồn/đích, sheet/cột, dòng bắt đầu, chính sách dữ liệu; không chứa secret |
| `tests/test_sync_nvl_stock.py` | Test thuật toán, tính toàn vẹn workbook và dữ liệu lỗi |
| `tests/test_nvl_publish_boundary.py` | Test mock Graph, chạy thử, xung đột và upload |
| `.github/workflows/sync-nvl-stock.yml` | Workflow riêng cho NVL |
| `README.md`, `.gitignore` | Hướng dẫn vận hành; bỏ qua output NVL cục bộ |

Ưu tiên không sửa hành vi helper chung; nếu cần tách module thì giữ API cũ và chạy regression. Không đổi global `sync_stock.DEST_PATH`, `DEST_SHEET`, `CODE_PATTERN` bằng monkey patch.

Giai đoạn đầu không cần state NVL commit lên Git: mỗi lần đọc hai workbook, lập diff và chỉ upload khi có ô cần đổi. Cách này tự xử lý mã mới ở đích, chỉnh tay D, và tránh xung đột commit state với workflow sản xuất. ETag vẫn phục vụ kiểm soát snapshot; không dùng riêng ETag nguồn để bỏ qua việc kiểm tra đích.

## 4. Quy tắc dữ liệu đề xuất cho phiên bản đầu

Các quy tắc ngoài mapping B/M → A/D là đề xuất triển khai, chưa phải quan sát từ workbook thật.

1. Có `source_start_row` và `target_start_row` được xác minh từ workbook; bỏ tiêu đề, dòng mã trống và dòng tổng được nhận diện rõ. Không coi mọi chuỗi trong A/B là mã vật tư.
2. Dùng hàm chuẩn hóa NVL riêng cho cả hai phía: trim khoảng trắng đầu/cuối; số nguyên và float nguyên tương đương; giữ mã text có số 0 đầu. Không tự xóa dấu, khoảng trắng bên trong, đổi hoa/thường hoặc bổ sung số 0. Mã số có format `000...` cần kiểm tra với mã text trước khi quyết định quy tắc, không tự khớp gần đúng.
3. Mã trùng sau chuẩn hóa trong phạm vi dữ liệu nguồn hoặc đích: báo rõ dòng và dừng publish. Không lấy dòng cuối hay tự cộng; chỉ bổ sung SUM khi nghiệp vụ xác nhận mỗi dòng là một kho/lô cần cộng.
4. Mã đích không có ở nguồn: giữ nguyên D và báo `missing_in_source`. Mã nguồn không có ở đích: báo `source_only`, không tự thêm dòng. Có thể cập nhật các mã khớp, nhưng báo trạng thái `completed_with_warnings` khi còn mã thiếu; không mô tả là đồng bộ đầy đủ.
5. Giá trị số 0 là tồn hợp lệ và phải ghi đè số cũ. Số âm/số lẻ giữ nguyên, có thể đánh dấu số âm trong report. Không tự làm tròn về nguyên hoặc chặn âm theo logic thành phẩm.
6. Với mã có ở đích: ô M trống, boolean, lỗi Excel, số không hữu hạn hoặc chuỗi không xác định được định dạng số là lỗi chặn publish. Không biến chúng thành 0. Chuỗi dấu phẩy/chấm chỉ parse theo quy ước được xác minh; không dùng phép xóa dấu phẩy một cách mặc định.
7. Nếu M là công thức, đọc giá trị cached và kiểm tra công thức bằng lượt đọc `data_only=False` hoặc XML. Thiếu cached value phải báo lỗi. Có cache không chứng minh cache còn mới; yêu cầu nguồn đã được Excel tính lại và lưu. Không chạy macro hay lưu đè nguồn.
8. Nếu D là công thức tại ô định cập nhật: báo lỗi cấu hình trước khi thay thế; xác minh D được dùng làm cột nhập tồn. Giữ các công thức khác, không xóa toàn sheet.
9. Không có mã đích hợp lệ hoặc không có mã nào khớp: dừng, không upload. Validation toàn bộ patch phải hoàn tất trước khi ghi file đích.

## 5. Thuật toán và bảo toàn workbook

```text
validate_config()
read stable source + target snapshots
target_rows = read valid Ton_NVL!A rows with original row numbers
stock = read Sheet1!B/M with NVL-specific validation
reject duplicate codes and invalid matched quantities
for each target row:
    if code is missing from stock: preserve D, report warning
    else if numeric D differs from stock[code]: add D-cell change
if no matched code: fail
patch only changed D cells in the original target ZIP
verify output values and preservation; emit proposal + report
if dry-run: stop without upload
if no changed cells: report unchanged, skip upload
recheck revisions and conditionally upload; verify remote result
```

- Lập dictionary một lần, duyệt nguồn và đích theo hàng: độ phức tạp O(n+m). Dùng `iter_rows()` thay vì đọc lại từng ô trên read-only worksheet. Xử lý dimension thiếu/sai để không bỏ sót dữ liệu.
- Tái sử dụng `find_sheet_xml_path`, `load_shared_strings`, `read_cell_text`, `set_numeric_cell` sau khi kiểm tra phù hợp. Resolve worksheet qua relationship, không hardcode `sheet1.xml`.
- Patch XML của đúng sheet `Ton_NVL`, chỉ đổi D ở các dòng khớp có thay đổi. Giữ style, cell ordering, hàng/cột, merge, validation, drawing, relationship và mọi ZIP part khác. Không round-trip cả workbook đích bằng `openpyxl.save()`.
- Nếu D chưa có cell XML, tạo đúng vị trí và cập nhật dimension khi cần; không làm mất định dạng sẵn có. D trong vùng merge/protection không tương thích phải báo lỗi trước upload.
- Kiểm thử công thức phụ thuộc D. Patch XML không tự tính công thức. Nếu cần yêu cầu Excel recalc qua `calcPr`, ghi rõ ngoại lệ cho `xl/workbook.xml`, kiểm thử riêng và không hứa cached kết quả đã được tính mới khi Excel chưa mở/tính lại.

## 6. Graph, retry và kết quả chạy

- Tái sử dụng `MS_TENANT_ID`, `MS_CLIENT_ID`, `MS_CLIENT_SECRET` hiện có. Xác minh app đọc được nguồn và ghi được đích; không in token/secret.
- Cấu hình file bằng path đã xác minh hoặc cặp drive ID/item ID. Kiểm tra tên và metadata để tránh nhầm file cùng tên. Không mặc định file đích nằm trong default drive nếu chưa xác minh.
- Đọc metadata trước/sau download; nếu ETag đổi thì tải lại để bytes và revision khớp nhau. Trước upload kiểm tra lại nguồn; thay đổi thì lập lại patch từ snapshot mới. Hai file không có giao dịch nguyên tử, cần nêu giới hạn nguồn có thể đổi sau kiểm tra cuối.
- Upload với ETag đích theo cơ chế hiện có và nghiệm thu khả năng chống stale write trên file test. HTTP 412: tải lại đích, đọc lại mapping A, lập lại patch, không gửi lại bytes cũ với ETag mới. 423/429/5xx/network: retry hữu hạn, tôn trọng `Retry-After`. 401/403/404 và lỗi dữ liệu: trả lỗi rõ ràng.
- Khi timeout làm kết quả upload không chắc chắn, đọc lại đích và so sánh trước khi upload tiếp. Xác minh sau upload; nếu có chỉnh sửa đồng thời, báo conflict thay vì tự rollback ghi đè người dùng.
- Report JSON gồm: source/target identity và revision; cấu hình; scanned/matched/changed/unchanged; missing/source-only/duplicates/errors; từng mã/dòng/ô/before/after; dry-run/published/unchanged/failed; số lần retry. `matched_count` và `changed_count` phải riêng.
- Output đề xuất: `nvl_stock_proposal.xlsx`, `nvl_stock_report.json`; lỗi vẫn xuất report. Dry-run chỉ ghi artifact local, không upload SharePoint và không sửa state sản xuất.

## 7. CLI và GitHub Actions

CLI dự kiến sau khi implement:

```bash
# Đọc SharePoint, chỉ tạo proposal
python -X utf8 sync_nvl_stock.py --config nvl_stock_config.json

# Đồng bộ SharePoint khi cấu hình và kiểm thử đã đạt
python -X utf8 sync_nvl_stock.py --config nvl_stock_config.json --publish

# Test bằng bản tải về, không cần credentials
python -X utf8 sync_nvl_stock.py --config nvl_stock_config.json --source-file "XNT_ketoan_Vikoda.xlsm" --target-file "Kế hoạch mua hàng.xlsx" --out offline_out/nvl
```

CLI từ chối kết hợp offline với `--publish`. Cấu hình thiếu target path/ID hoặc dòng bắt đầu hợp lệ phải lỗi rõ, không tự đoán.

Workflow riêng: `workflow_dispatch` có boolean `publish`, mặc định false; test gate giống repo; secrets chỉ cấp cho bước online; upload artifact với `if: always()` và retention 14 ngày. Không cần `contents: write` vì không commit state; dùng `contents: read`. Concurrency `sharepoint-nvl-stock-sync`, `cancel-in-progress: false`.

Sau nghiệm thu, bổ sung lịch 06:10 Việt Nam (`10 23 * * *` UTC) hoặc giờ nghiệp vụ yêu cầu. Đây là lịch đề xuất, chưa được người dùng chốt. Dùng event `sharepoint_nvl_stock_updated` từ Power Automate theo dõi file nguồn; cấu hình để thay đổi file đích không kích hoạt vòng lặp. Nếu muốn tái sử dụng event cũ, ghi rõ nó sẽ đánh thức cả hai workflow. Chỉ bật auto-publish sau khi chạy thử với dữ liệu thật đạt yêu cầu.

## 8. Kiểm thử và nghiệm thu

1. Nguồn/đích khác thứ tự dòng: D vẫn đúng theo mã; cột M không bị đổi đơn vị.
2. Mã số/text, float nguyên, khoảng trắng, mã chữ, số 0 đầu; tiêu đề/tổng không được coi là vật tư.
3. Số 0, âm, lẻ; blank/error/bool/NaN/infinity; chuỗi số mơ hồ; công thức có/thiếu cache.
4. Mã trùng nguồn/đích; mã thiếu nguồn giữ D và có warning; mã chỉ có nguồn không thêm dòng; không khớp mã nào thì không upload.
5. Thiếu sheet, sai header, dimension thiếu/sai, D chưa có cell, D là công thức/merge; không chỉnh nhầm sheet.
6. So sánh bytes nội dung từng ZIP part ngoài phần cho phép; kiểm tra giá trị/công thức/style và XML ngoài các cell D trong sheet đích.
7. Chạy hai lần với cùng dữ liệu: lần hai `changed_count=0`, không gọi upload; thêm mã đích hoặc sửa D vẫn được phát hiện khi ETag nguồn không đổi.
8. Mock Graph: dry-run không upload; validation lỗi không upload; 412 tải lại và giữ sửa đổi khác của người dùng; 429 tuân thủ Retry-After; retry hữu hạn; timeout sau upload không gây ghi lại vô ích; nguồn đổi giữa download/publish thì tính lại.
9. Toàn bộ test hiện có phải pass: `python -X utf8 -m unittest discover -s tests -v`.
10. Chạy offline/dry-run dữ liệu thật: đối chiếu tối thiểu 10 mã có tồn dương/0/âm/lẻ nếu có; kiểm tra danh sách missing/duplicate; mở proposal bằng Excel và xác nhận không có repair prompt. Publish trên bản sao để nghiệm thu ETag và công thức phụ thuộc trước lần chạy chính thức.

## 9. Trình tự giao việc cho Antigravity

1. Đọc `AGENTS.md` nếu có và các file phân tích ở mục 2; cập nhật từ `origin/main` mới nhất. Tạo nhánh tính năng, không ghi đè thay đổi local của người dùng.
2. Xác minh cấu trúc workbook và vị trí đích; hoàn thiện config. Có thể implement và test bằng fixture trước khi có kết nối thật, nhưng không tự điền đường dẫn hoặc dòng dữ liệu suy đoán để publish.
3. Implement module và CLI, rồi test offline/mock Graph theo các tiêu chí trên.
4. Thêm workflow chạy thủ công, tài liệu và gitignore; chạy toàn bộ test, tạo proposal đối chiếu dữ liệu thật khi truy cập được.
5. Commit/push nhánh tính năng lên `Huanpro1239/Planning_vikoda`, tạo PR mô tả mapping B/M → A/D, thay đổi, test và cấu hình còn thiếu. Không đưa workbook nghiệp vụ, token hoặc báo cáo chứa dữ liệu thật vào Git.
6. Bàn giao URL PR, kết quả test, artifact chạy thử, danh sách cấu hình đã xác minh/chưa xác minh. Chỉ ghi nhận hoàn tất live sync sau khi upload và kiểm tra được file đích. Bật lịch/event theo cấu hình vận hành đã chốt.

Prompt dùng kèm: **“Hãy triển khai theo PLAN_TON_NVL_ANTIGRAVITY.md. Yêu cầu đã chốt: XNT_ketoan_Vikoda.xlsm/Sheet1 cột B là mã, cột M là tồn; Kế hoạch mua hàng.xlsx/Ton_NVL cột A là mã, cột D nhận tồn. Tái sử dụng kết nối SharePoint của repo, thêm luồng NVL riêng, hoàn thành test và PR GitHub theo kế hoạch. Phân biệt rõ cấu hình đã xác minh với giả định; chưa đủ dữ liệu thật thì bàn giao phần code/test hoàn chỉnh và nêu phần chưa nghiệm thu.”**

## Tài liệu kỹ thuật đối chiếu

- [Microsoft Graph — upload nội dung file](https://learn.microsoft.com/en-us/graph/api/driveitem-put-content?view=graph-rest-1.0).
- [Microsoft Graph — upload session và If-Match](https://learn.microsoft.com/en-us/graph/api/driveitem-createuploadsession?view=graph-rest-1.0): phương án có điều kiện nếu cần điều chỉnh cơ chế upload sau nghiệm thu.
- [openpyxl — data_only và giới hạn bảo toàn workbook](https://openpyxl.readthedocs.io/en/stable/tutorial.html): cached value là giá trị lưu sẵn, không phải kết quả tính mới của lần đồng bộ.
- [openpyxl — read-only và dimension](https://openpyxl.readthedocs.io/en/stable/optimized.html).
