# Review vòng 5 — dry-run dữ liệu thật

Ngày 09/09/2026. PR #22: https://github.com/Huanpro1239/Planning_vikoda/pull/22
Commit đã kiểm tra: `4036fe3` (nhánh `feat/sync-nvl-stock`).
Worktree: `D:\Vikoda\Planning_vikoda_nvl`.

## Kết luận

Hai điểm tồn của vòng 4 đã sửa đúng. Mapping nghiệp vụ B→A, M→D đã được đối chiếu với file thật và đúng.
**Chưa thể merge**: `scripts/survey_nvl_sharepoint.py` chắc chắn crash và đang được gắn vào workflow sản xuất
`sync-stock.yml`, làm hỏng bước lưu state sau khi publish kế hoạch. Bản thân module `sync_nvl_stock.py` không có lỗi chặn.

## Đã đạt (có bằng chứng)

- `python -X utf8 -m unittest discover -s tests`: **165/165 PASS** (20.9s).
- Vòng 4 mục 1: `upload_acknowledged = not is_timeout_upload`, tách `upload_may_have_committed`; 3 test timeout đã assert đúng.
- Vòng 4 mục 2: `_safe_close_workbook` đã đóng thêm `vba_archive`.
- Mapping thật (đọc từ `dry_run_artifacts/dry_run_artifacts/`):
  - Nguồn `Sheet1`: dòng 9/10 là header ghép `Tồn cuối kỳ` → **M = Số lượng**, N = Thành tiền. Dữ liệu từ **dòng 11** đến 366, 356 mã, **không có ô M rỗng**. `start_row=11` đúng.
  - Đích `Ton_NVL`: header `A=Mã NVL, B=Tên NVL, C=ĐVT, D=Tồn Cuối, E=Tồn Đơn hàng`, dữ liệu từ dòng 2, 144 mã. `start_row=2`, cột A/D đúng.
  - Đối soát: 137 khớp (toàn bộ D đang rỗng → 137 ô cần ghi), 7 mã đích không có trong nguồn (giữ nguyên, đã có warning), 219 mã nguồn không có ở đích (bỏ qua, đúng thiết kế).
  - **Không có công thức nào trong file đích tham chiếu `Ton_NVL`** (0 hit, 0 defined name) → không có rủi ro recalc ở thời điểm này.
  - Proposal: các part ZIP khác giữ nguyên byte, `dimension` vẫn `A1:E145`, dòng `missing_in_source` (vd. dòng 51) không sinh ô D.

## Chặn merge

### 1. `scripts/survey_nvl_sharepoint.py` crash — và kéo theo hỏng workflow sản xuất

`NVLReconcileResult` không có `total_rows`, `duplicate_codes_source`, `duplicate_codes_target`, nhưng script đọc cả ba khi
dựng `audit_summary` (`scripts/survey_nvl_sharepoint.py:224-231`). Đã tái hiện bằng chính hai file thật đã tải về:

```
source codes 356 changes 137 unchanged 0 missing 7
REPRO CRASH: AttributeError 'NVLReconcileResult' object has no attribute 'total_rows'
```

Bằng chứng lần chạy CI thật: artifact tải về có `nvl_stock_proposal.xlsx`, `nvl_stock_report.json`, hai file workbook thật,
nhưng **không có `audit_summary.json`** — đúng điểm crash.

Hệ quả nghiêm trọng: bước `Survey and dry-run NVL stock` được chèn vào `.github/workflows/sync-stock.yml:90`,
tức workflow kế hoạch sản xuất chạy cron 06:00 hằng ngày và nhận `repository_dispatch`. Bước này nằm **sau** bước publish
và **trước** bước `Save sync states after authorized publish` (`if: success()`). Khi survey crash:
kế hoạch đã publish lên SharePoint nhưng `state.json` / `planning_runtime.json` không được commit
→ fingerprint chống lặp lệch với thực tế → lần chạy sau tính lại và publish lại.

Xử lý: bỏ hẳn bước NVL khỏi `sync-stock.yml` (đã có workflow riêng `sync-nvl-stock.yml`). Nếu vẫn muốn giữ,
tối thiểu phải `continue-on-error: true` và đặt sau bước lưu state — nhưng khuyến nghị là tách hẳn.

### 2. Script survey tự ghi đè `nvl_stock_config.json` bằng suy đoán

`scripts/survey_nvl_sharepoint.py:146-157` ghi lại `target.sharepoint_path` và `source.start_row` vào file config:

- `target.sharepoint_path` lấy từ **file đầu tiên** trùng tên khi duyệt BFS tới độ sâu 3 — có thể trúng bản sao/lưu trữ.
  Đã có sẵn `target.sourcedoc = 89D1BA7B-006E-4527-B879-ABF120309214` nhưng **không hề đối chiếu**.
- `source.start_row` = dòng đầu tiên trong 1..15 có M là số; không kiểm tra B có mã hay không.
  Ở file hiện tại ra 11 (đúng) chỉ vì may mắn — một con số lạc ở M9/M10 là lệch toàn bộ dữ liệu.
- Trong CI thì thay đổi này bị vứt đi (runner ephemeral) → không có lợi ích, chỉ có rủi ro.

Xử lý: chốt cứng path/start_row trong config đã review; script survey chỉ **đọc và báo cáo**, và phải fail rõ ràng
nếu item tìm được có `sourcedoc`/id khác với config.

### 3. Artifact CI chứa nguyên workbook kế toán thật

`sync-stock.yml:107` upload cả thư mục `dry_run_artifacts/`, trong đó có `real_source_XNT_ketoan_Vikoda.xlsm`
(báo cáo nhập xuất tồn kế toán, có giá trị tiền) và `real_target_Ke_hoach_mua_hang.xlsx`. Artifact giữ 14 ngày và
ai có quyền đọc repo đều tải được. Chỉ nên upload `nvl_stock_report.json` + proposal, không upload file nguồn thật.

## Rủi ro nghiệp vụ cần chốt trước khi publish

1. **Kỳ dữ liệu nguồn**: `Sheet1!A7` ghi `Từ ngày 01-08-2026 đến ngày 31-08-2026`, tức tồn cuối **tháng 8** trong khi chạy ngày 09/09.
   Module không đọc và không kiểm tra kỳ. File kế toán không được refresh sẽ được đẩy thẳng vào `Tồn Cuối` mà không cảnh báo.
   Đề xuất: đọc dòng kỳ báo cáo, ghi vào report, và cảnh báo/chặn khi kỳ không bao phủ ngày chạy.
2. **Đơn vị tính**: đích có cột `ĐVT` (Kg…), nguồn là `Số lượng` của `Tồn cuối kỳ`; không có bước đối chiếu đơn vị. Cần business ký nhận.
3. **7 mã đích không có trong nguồn** (330500109, 330500110, …): hiện giữ nguyên giá trị cũ (đang rỗng). Cần xác nhận đây là mã mới chưa phát sinh hay là mã sai.

## Điểm nhỏ nên sửa kèm

1. `direct_copy` và `preserve_missing_in_source` trong config **không được đọc ở bất kỳ đâu**. Đặt `preserve_missing_in_source: false`
   sẽ không có tác dụng gì. Hoặc thực thi, hoặc bỏ khỏi config.
2. `parse_nvl_quantity` làm hỏng **toàn bộ** lần chạy nếu một dòng nguồn có mã hợp lệ nhưng M rỗng/không phải số
   (`sync_nvl_stock.py:480`). Hiện file thật không có ô rỗng nào, nhưng một ô trống trong bản xuất kế toán sẽ chặn cả ngày đồng bộ.
   Cân nhắc bỏ qua + warning theo dòng, hoặc thêm policy flag.
3. Patch không cập nhật thuộc tính `spans` của `<row>`: **122 dòng** sau patch mang `spans="1:3"` nhưng đã có ô `D`.
   `spans` chỉ là gợi ý nên Excel thường tự tính lại, nhưng nên set lại cho đúng chuẩn.
4. `verify_nvl_patched_workbook` kiểm tra ô cũ không bị đổi/mất, nhưng **không kiểm tra ô mới phát sinh ngoài `changed_refs`**
   (`patch_cells - orig_cells - changed_refs`). Thêm assertion này là đóng kín được vòng xác minh.
5. Không có test nào chạm vào `scripts/` — chính vì vậy lỗi mục 1 lọt qua cổng test 165/165.
6. `scripts/trigger_and_download_dryrun.py` là công cụ dev cục bộ (hardcode repo/branch, lấy token từ git credential).
   Nên để trong `temp_audit/` (đã gitignore) thay vì commit vào `scripts/`.

## Bước tiếp theo

1. Sửa mục 1–3 phần chặn merge. Chạy lại 165 test.
2. Chốt kỳ dữ liệu và đơn vị tính với kế toán/mua hàng; bổ sung kiểm tra kỳ vào report.
3. Chạy lại dry-run trên workflow `sync-nvl-stock.yml` (không phải `sync-stock.yml`), bàn giao proposal + report.
4. Mở proposal bằng Excel thật xác nhận không hiện repair, rồi mới thử publish trên bản sao. Sau đó mới merge và chạy chính thức.
