# Review vòng 2 — PR #22

Ngày 09/09/2026. Head: `4878ea3b3f7e4768703d220d84ed9349feec87d0`.
PR: https://github.com/Huanpro1239/Planning_vikoda/pull/22

**Kết luận: phần xử lý đã cải thiện, nhưng chưa đạt để publish chính thức.** Chưa có nghiệm thu workbook SharePoint thật; target.sharepoint_path vẫn để trống. Chỉ chạy kiểm thử local và mock Graph, không ghi SharePoint.

## Đã xác nhận sửa đúng

- Chạy toàn bộ suite: 148/148 PASS, 13.650 giây.
- Probe lần trước: `1,5` đọc thành 1.5; `1.234,56` thành 1234.56; `1,2,3` bị từ chối.
- Dimension thiếu/ngắn vẫn đọc đủ hai mã VT001/VT002.
- Giữ nguyên D='-' ở mã thiếu nguồn, verify không crash.
- Nguồn đổi 100 → 999 sau download được tải lại và upload 999.
- Upload 403 ghi report failed. Test retry download và timeout đã có trong suite.

## 1. [P1] Không được báo published khi xác minh sau upload thất bại

Vị trí: `sync_nvl_stock.py:1215-1239`, khối `post_upload_verify` và `finalize_published`.

Đã tái hiện bằng mock: upload trả thành công, nhưng dữ liệu tải lại có D2=777 thay vì 100. Hàm xác minh phát hiện sai và ném lỗi; catch chỉ print cảnh báo rồi tiếp tục. Report cuối có `status='published'`, `warnings=[]`, hàm trả thành công. Đây có thể là chỉnh sửa đồng thời hoặc kết quả đọc lại chưa đúng; chương trình không thể kết luận đã nghiệm thu thành công.

Yêu cầu sửa:

- Chỉ đánh dấu published khi xác minh kết quả đạt.
- Lỗi GET tạm thời sau upload: retry bước đọc/xác minh có giới hạn, không tự upload lại toàn bộ.
- Đích khác patch hoặc nội dung khác đã bị sửa đồng thời: ghi trạng thái failed/conflict/verification_failed và phase post_upload_verify; giữ thông tin upload đã nhận phản hồi thành công, tránh nói chắc chưa ghi gì. Không tự rollback hoặc ghi đè lại thay đổi của người khác.
- Xác minh hết lượt vẫn không đọc được: dùng trạng thái rõ chưa xác minh được, không published; báo lỗi qua CLI.
- Thêm test D bị đổi sau upload, ô ngoài D bị sửa đồng thời, GET sau upload lỗi tạm thời rồi phục hồi, và hết retry. Các test phải kiểm tra trạng thái report, số lần upload, kết quả CLI/hàm và dữ liệu giữ nguyên.

## 2. [P2] Một số nhánh lỗi không xuất error report

Vị trí: các nhánh source-freshness dùng continue trong `run_nvl_sync`, lệnh raise cuối vòng lặp tại `sync_nvl_stock.py:1281`; kiểm tra config trước try tại khoảng `:1088`.

Đã tái hiện: nguồn đổi mỗi lần download, hết 2 lượt thì RuntimeError và không tồn tại nvl_stock_report.json. Config thiếu target path cũng không tạo report. Nếu lượt trước đã sinh proposal rồi source đổi trước upload, có thể còn proposal mà không có error report giải thích vì sao không publish.

Yêu cầu sửa:

- Đưa các đường thoát lỗi vào cơ chế finalize chung để ghi report failed với phase, attempt, revisions và lý do trước khi ném lỗi.
- Bao gồm hết lượt freshness, cấu hình thiếu/sai, lỗi load config và lỗi offline input nếu thư mục output ghi được.
- Quản lý proposal còn lại rõ ràng: xóa artifact không còn hợp lệ hoặc đánh dấu invalid trong report, không để bị hiểu nhầm là proposal đạt.
- Test nguồn đổi liên tục cả sau download lẫn trước upload; xác nhận có error report và không upload.

## Những phần phải nghiệm thu tiếp

- Xác minh target.sharepoint_path, đúng workbook/drive, dòng bắt đầu, header, dạng mã và công thức D; không đoán path từ sourcedoc GUID.
- Chạy dry-run với file thật. So sánh ít nhất 10 mã và các trường hợp 0/âm/lẻ nếu có; báo các mã thiếu/trùng và kiểm tra proposal bằng Excel.
- Workflow hiện chỉ chạy thủ công, chưa bật tự động. Chưa publish file chính thức.

## Bằng chứng

- `temp_audit/nvl_rereview_tests.log`: 148 tests PASS.
- `temp_audit/nvl_review_probe.py` / `.json`: các ca lỗi vòng trước đã được chạy lại.
- `temp_audit/nvl_rereview_probe.py` / `.json`: tái hiện hai lỗi vòng này.

## Prompt giao Antigravity

Đọc REVIEW_PR22_ROUND2.md. Trên nhánh feat/sync-nvl-stock, sửa hai lỗi còn lại: không báo published khi post-upload verification thất bại; luôn xuất error report khi freshness retry hết lượt hoặc config/input lỗi. Không tự upload lại/rollback để sửa conflict sau upload. Bổ sung regression tests, chạy toàn bộ suite và hai probe review, rồi push cập nhật vào PR #22. Cập nhật PR description đúng kết quả nghiệm thu. Sau đó xác minh đường dẫn/cấu trúc SharePoint và chạy dry-run dữ liệu thật nếu có quyền, bàn giao proposal/report cùng đối chiếu ít nhất 10 mã. Nếu không truy cập được, báo chính xác thông tin còn thiếu. Chưa merge, chưa publish workbook chính thức, chưa bật lịch. Báo commit mới, test và phần còn chưa nghiệm thu.
