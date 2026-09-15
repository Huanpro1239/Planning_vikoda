# Review commit 6d48041 — nghiệm thu dry-run

Ngày 09/09/2026. Đã kiểm tra code, chạy test local và đọc trạng thái GitHub Actions qua API chỉ đọc.

## Kết quả

- 170/170 tests PASS, 22.957 giây (`temp_audit/round6_tests.log`).
- Workflow NVL riêng run 34321537269, SHA `6d4804179f4f7814e0a5de92ee800d77e2420bb7`, completed/success: https://github.com/Huanpro1239/Planning_vikoda/actions/runs/34321537269 . PR test cùng SHA cũng success.
- Đã bỏ survey NVL khỏi workflow sản xuất. Survey hoàn tất, có audit_summary.json, không còn lỗi thuộc tính thiếu.
- Config nguồn dòng 11, đích dòng 2, B/M → A/D đúng với snapshot thật. Audit ghi đúng ID và ETag chứa GUID của nguồn/đích được chỉ định.
- Proposal mới: đối chiếu độc lập 137/137 mã đúng giá trị nguồn; 137 ô D thay đổi; 7 mã thiếu nguồn giữ trống. Chỉ XML sheet Ton_NVL thay đổi giữa ZIP nguồn đích và proposal. Bằng chứng `temp_audit/round6_real_audit.json`.

## Chưa chốt để publish chính thức

Nguồn hiện là kỳ 01–31/08/2026. Phải xác nhận sử dụng số chốt kỳ này hay cần file mới hơn; không suy ra ngày sửa file là ngày chốt tồn.

Audit có 95 dòng khác chuỗi đơn vị:

- 74 dòng `CAI` → `Cái`: khác cách viết, chưa phải bằng chứng khác đại lượng.
- 17 dòng đơn vị đích trống: 11 KG và 6 CAI.
- 4 mã cần xác minh cụ thể:
  - 430100118, dòng 81: BIH → Cái.
  - 430100126, dòng 83: CAI → Bình.
  - 430200173, dòng 126: CAI → Kg (nhãn thân Vikoda PET 1.5 Lít PVC). Đây là chênh lệch có thể làm sai ý nghĩa số tồn khi chép trực tiếp.
  - 430400002, dòng 144: CN → Cái.

Không tự đổi đơn vị, quy đổi số lượng hoặc sửa cột C. Cần xác minh đây là nhãn đơn vị sai/viết tắt hay đơn vị nghiệp vụ khác. Các giá trị M → D đã đúng về mặt số học.

Chưa có bằng chứng mở proposal bằng Excel thật không repair, chưa nghiệm thu upload trên bản sao. Dry-run hiện chứng minh đọc và tạo proposal; không chứng minh việc ghi SharePoint đã hoạt động.

## Điểm nhỏ còn lại ở công cụ hỗ trợ

`scripts/trigger_and_download_dryrun.py` đã kiểm tra conclusion, nhưng chọn run chỉ theo tên/nhánh, không dùng start_time hoặc expected SHA. Dispatch thất bại có thể chọn run thành công cũ. Sửa để gắn đúng SHA/run mới hoặc yêu cầu run ID rõ ràng; không dùng artifact cũ làm chứng cứ cho code mới. Lần này reviewer đã kiểm tra độc lập run đúng SHA nên kết quả dry-run nêu trên vẫn có giá trị.

`scripts/survey_nvl_sharepoint.py` chỉ in cảnh báo khi identity khác rồi vẫn chạy. Nên đưa mismatch vào audit và dừng nghiệm thu thay vì status success khi sai file. Trong artifact đang review, cả hai GUID đều khớp, nên không có bằng chứng chọn sai file ở lần này.

## Prompt tiếp theo

Đọc REVIEW_6d48041.md. Code và dry-run đã qua review; chuyển sang chốt dữ liệu và nghiệm thu publish trên bản sao. Lập danh sách 4 mã khác đơn vị cần xác minh, 17 dòng thiếu ĐVT và 7 mã thiếu nguồn. Chuẩn hóa cách so sánh CAI/Cái trong báo cáo bằng alias rõ ràng, không sửa dữ liệu workbook hoặc tự quy đổi. Ghi rõ nguồn kỳ 01–31/08/2026 để người dùng chốt. Sửa helper chọn workflow run theo đúng SHA/run ID và fail rõ khi survey identity mismatch. Chạy regression tests và dry-run lại nếu code thay đổi. Bàn giao proposal, report, audit và bằng chứng Excel mở không repair nếu thực hiện được. Chưa publish workbook chính thức/chưa bật lịch; chờ chốt kỳ và đơn vị trước bước đó.
