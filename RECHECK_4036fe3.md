# Kết quả kiểm tra độc lập commit 4036fe3

Ngày 09/09/2026. PR #22 vẫn mở, head `4036fe3f27d9b41c0a80c0ee1904222ff5b8ff0a` tại thời điểm đọc metadata. Đây là review code và các snapshot đã tải về, không phải nghiệm thu ghi SharePoint trực tiếp.

## Đã đạt

- Toàn bộ suite chạy lại: 165/165 PASS (21.544 giây), `temp_audit/round5_tests.log`.
- Cấu hình đã điền target path `Tinh san xuat Mua hang 2027/Kế hoạch mua hàng.xlsx` và source_start_row=11. Nội dung snapshot xác nhận dòng 9/10 nguồn là header, dòng 11 bắt đầu mã; đích dòng 1 có `Mã NVL`, `Tên NVL`, `ĐVT`, `Tồn Cuối`.
- Đối chiếu độc lập trực tiếp từ workbook, không dùng parser/mapping của module: 137/137 mã khớp M nguồn với D proposal trong dung sai 1e-6; không có sai lệch. Có 1 số 0 và 52 số lẻ, không có tồn âm trong 137 mã này.
- Chỉ ZIP part `xl/worksheets/sheet3.xml` thay đổi; đối chiếu cell XML có đúng 137 cell thay đổi, tất cả ở D. Các part khác giữ nguyên.
- 7 mã không có nguồn giữ nguyên D trống: 330500109, 330500110, 330500111, 330500112, 430100018, 430200147, 430200154.
- Bằng chứng số liệu và 10 mã mẫu: `temp_audit/round5_real_audit.json`; script read-only `temp_audit/round5_real_audit.py`, chạy bằng bundled Python với `-X utf8`.

## Hai việc code/workflow cần hoàn tất trước merge

### [P1] Script survey gọi thuộc tính không tồn tại và có thể chặn lưu state sản xuất

`scripts/survey_nvl_sharepoint.py:224,229,230` gọi `rec.total_rows`, `rec.duplicate_codes_source`, `rec.duplicate_codes_target`. Kiểm tra trực tiếp `NVLReconcileResult` cho thấy cả ba đều không tồn tại. Khi đến bước tạo audit_summary sau khi tạo proposal, script sẽ ném AttributeError. Artifact hiện có proposal/report nhưng không có audit_summary phù hợp với đường lỗi này; chưa đọc được log GitHub run nên không khẳng định kết luận của run từ artifact đơn lẻ.

`.github/workflows/sync-stock.yml:90` thêm survey không có điều kiện vào workflow sản xuất, nằm sau pipeline publish và trước bước Save sync states (`if: success()`). Nếu merge như hiện tại, lỗi survey có thể làm bước commit state bị bỏ qua dù workbook sản xuất đã publish. Đây là ảnh hưởng ngoài phạm vi đồng bộ NVL.

Sửa: bỏ bước survey và artifact NVL khỏi workflow sản xuất; giữ chạy NVL ở workflow riêng. Dựng audit từ thuộc tính thật (`len(rec.target_codes)` và số liệu validation đã kiểm tra), không bịa số duplicate. Thêm test chạy script survey tới cuối với mock Graph/snapshot, assert có audit_summary và exit 0. Khi lỗi phải exit khác 0 và có error report.

### [P2] Công cụ tải artifact báo hoàn tất ngay cả khi workflow thất bại

`scripts/trigger_and_download_dryrun.py` đọc conclusion nhưng không kiểm tra trước khi kết thúc thành công; artifact có thể được tạo trước lỗi và upload qua `if: always()`. Không được coi có artifact là workflow đã chạy đạt.

Sửa: có thể tải artifact lỗi để chẩn đoán, nhưng phải báo conclusion/run ID/head SHA và trả exit code lỗi nếu conclusion khác success. Chạy lại survey riêng tới cuối sau khi sửa, bàn giao run thành công và audit_summary. Không trigger pipeline sản xuất chỉ để khảo sát NVL.

## Việc dữ liệu/vận hành còn cần

- Đọc trực tiếp `Sheet1!A7` nguồn: `Từ ngày 01-08-2026 đến ngày 31-08-2026`. Chốt người dùng muốn lấy tồn cuối tháng 8 hay tồn mới hơn trước lần publish. Ghi kỳ nguồn vào report; không tự đặt quy tắc kỳ phải bao phủ ngày chạy vì nghiệp vụ có thể dùng số chốt tháng trước.
- Nguồn có cột F là Đvt; đối chiếu với cột C đích và báo sai lệch trước khi chép trực tiếp. Không tự quy đổi.
- 7 mã thiếu nguồn giữ nguyên theo quy tắc đã đặt; báo rõ chưa có số liệu, không coi là tồn 0.
- Script survey hiện tìm file đầu tiên trùng tên và tự ghi config. Đã có cấu hình path cụ thể thì dùng cấu hình đó, xác minh identity với metadata SharePoint; survey nên báo đề xuất thay vì âm thầm thay cấu hình chạy thật.
- Chưa mở proposal trong Excel thật để xác nhận không repair và chưa thử upload trên bản sao. Việc đọc thành công bằng openpyxl không thay thế kiểm tra này.
- Không cần sửa parser để tự bỏ qua ô M lỗi ở mã khớp: chính sách dừng với dữ liệu không hợp lệ là có chủ ý trong kế hoạch gốc. Tránh mở rộng tính năng trong đợt nghiệm thu.

## Prompt giao Antigravity

Đọc RECHECK_4036fe3.md. Mapping dữ liệu thật đã được review đạt 137/137 mã. Hoàn tất phần vận hành: sửa survey dùng đúng thuộc tính NVLReconcileResult và test chạy đến cuối có audit_summary; bỏ survey/artifact NVL khỏi workflow sản xuất sync-stock.yml, dùng workflow NVL riêng; sửa helper tải artifact để không báo thành công khi workflow thất bại. Dùng target path đã cấu hình, xác minh identity, ghi kỳ nguồn và đối chiếu đơn vị nguồn F với đích C vào audit. Chạy lại dry-run riêng, bàn giao run ID/head SHA/conclusion, proposal, report, audit_summary và 7 mã thiếu nguồn. Nguồn hiện là kỳ tháng 8/2026: nêu rõ để người dùng chốt kỳ, không tự suy đoán. Cập nhật PR description đang còn thông tin target path trống cũ. Chưa merge, chưa publish file chính thức, chưa bật lịch. Không sửa nhánh KHSX_ki.
