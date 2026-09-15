# Review vòng 4 — sẵn sàng nghiệm thu dry-run

Ngày 09/09/2026. PR #22: https://github.com/Huanpro1239/Planning_vikoda/pull/22
Commit đã kiểm tra: `b3f4db57a50c9d4041f61cc69f943aea743747c9`.
Worktree: `D:\Vikoda\Planning_vikoda_nvl`.

## Kết luận

Hai lỗi chặn của vòng 3 đã được sửa và kiểm chứng lại. Có thể chuyển sang dry-run dữ liệu thật. Chưa đủ bằng chứng kết luận vận hành SharePoint chính thức đạt: target.sharepoint_path còn trống, chưa đối chiếu cấu trúc/dữ liệu workbook thật và chưa kiểm thử upload trên bản sao.

## Bằng chứng mới

- `python -X utf8 -m unittest discover -s tests -v`: 165/165 PASS, 15.641 giây; log `nvl_review_round4_tests.log`.
- Probe vòng 3: timeout sau server commit và chỉnh sửa D=777 → chỉ 1 lần upload, D vẫn 777, report failed; không còn PUT lần 2 ghi đè.
- Probe vòng 3: config không tồn tại sau lần chạy thành công → CLI exit 1, report failed/init_cli, proposal cũ bị xóa.
- Probe vòng 1 và vòng 2 đã chạy lại; kết quả được lưu trong `temp_audit/round4_round1.log`, `temp_audit/round4_round2.log` và các JSON probe.
- Tất cả kiểm tra trên dùng fixture/mock Graph; không ghi workbook SharePoint thật. Không sửa code sản phẩm.

## Hai điểm nhỏ nên chỉnh kèm nghiệm thu

1. Trong error report sau timeout, code vẫn gán `upload_acknowledged=True` dù chưa nhận ACK. Có `is_timeout_upload=True` và trạng thái failed nên không còn lỗi báo published, nhưng trường ACK gây hiểu nhầm. Đặt ACK false/unknown khi timeout; tách `upload_may_have_committed` khỏi việc nhận phản hồi. Bổ sung assertion vào test hiện có, không cần thay thuật toán.
2. Các probe vẫn phát sinh cảnh báo dọn `ZipFile.__del__` trên Python 3.12 (`I/O operation on closed file`), dù exit code 0 và kết quả đối soát đạt. Kiểm tra vòng đời stream/vba_archive và fixture để đóng đúng tài nguyên; không chỉ ẩn cảnh báo. Đây chưa phải bằng chứng workbook bị hỏng.

## Bước tiếp theo

1. Làm việc trong đúng worktree NVL; giữ nhánh KHSX_ki riêng.
2. Xác minh đường dẫn/drive của `Kế hoạch mua hàng.xlsx` từ quyền truy cập hiện có hoặc bản tải về. Không dùng sourcedoc GUID làm Graph item ID, không đoán thư mục. Xác nhận Sheet1 B/M và Ton_NVL A/D, dòng bắt đầu dữ liệu, mã, đơn vị và công thức D.
3. Hoàn thiện config và chạy offline hoặc online dry-run (không có --publish). Nếu chỉ có secrets trên GitHub, chuẩn bị đường chạy dry-run trên nhánh PR phù hợp; không tự bật production publish chỉ để thử kết nối.
4. Bàn giao proposal/report, bảng đối chiếu ít nhất 10 mã (0/âm/lẻ nếu có), số mã khớp/thiếu/trùng, xác nhận Excel mở proposal không repair và công thức phụ thuộc D tính đúng sau recalc. Chưa xác minh được điều gì phải ghi rõ.
5. Khi code/config và dry-run thật đạt, mới quyết định merge và thử publish trên bản sao. Sau đó mới chạy chính thức và chốt lịch tự động. Hiện workflow vẫn chỉ chạy thủ công.

## Prompt cho Antigravity

Đọc REVIEW_PR22_ROUND4.md trong D:\Vikoda\Planning_vikoda_nvl. Hai lỗi vòng 3 đã được review đạt trên commit b3f4db5. Chuyển sang nghiệm thu dữ liệu thật: xác minh đường dẫn workbook đích, dòng bắt đầu, mapping B/M → A/D, hoàn thiện config và chạy dry-run. Bàn giao proposal/report cùng đối chiếu ít nhất 10 mã, danh sách mã thiếu/trùng và kết quả kiểm tra Excel. Chỉnh trường upload_acknowledged cho đúng khi timeout và kiểm tra cảnh báo đóng ZipFile theo review, giữ test regression pass. Nếu chưa truy cập được SharePoint, báo chính xác thông tin cần bổ sung hoặc dùng bản tải về để nghiệm thu offline. Push cập nhật PR #22, báo commit và kết quả. Chưa publish workbook chính thức, chưa bật lịch; chưa tự merge trước khi dry-run thật được kiểm tra.
