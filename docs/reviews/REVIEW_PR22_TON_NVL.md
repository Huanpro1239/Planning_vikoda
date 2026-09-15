# Review PR #22 — đồng bộ Ton_NVL

Ngày: 09/09/2026. PR: https://github.com/Huanpro1239/Planning_vikoda/pull/22
Head đã review: `5cd24be8e2e0607860e64023554a6e7cde978190`. PR đang mở, chưa merge tại thời điểm kiểm tra.

**Kết luận: cần sửa trước khi merge và publish dữ liệu thật.** Mapping B/M → A/D và kiến trúc luồng NVL riêng đúng. Bộ test hiện có chạy lại đạt 137/137, nhưng các ca bổ sung dưới đây tái hiện lỗi chưa được bộ test bao phủ. Chưa kiểm tra hai workbook SharePoint thật; mọi probe dùng workbook tổng hợp và mock Graph, không upload lên SharePoint.

## Các lỗi cần sửa

### 1. [P1] Parser tự bỏ dấu phẩy làm sai số tồn

Vị trí: `sync_nvl_stock.py:196`.

Đã tái hiện `parse_nvl_quantity('1,5') == 15.0`, `'1.234,56' == 1.23456`, và chuỗi không hợp lệ `'1,2,3' == 123.0`. Nếu nguồn có số lưu dạng text với dấu phẩy thập phân, giá trị ghi vào D sai mà không báo lỗi. Chưa có xác nhận workbook thật dùng quy ước số nào.

Sửa: số Excel dạng numeric được chép trực tiếp; chuỗi phải tuân thủ định dạng được cấu hình/xác minh, kiểm tra dấu phân nhóm nghiêm ngặt. Khi chưa có quy ước, chặn chuỗi mơ hồ. Thêm test số Việt Nam, số kiểu Anh, grouping sai và lỗi không upload.

### 2. [P1] Reader phụ thuộc dimension, có thể bỏ sót dòng mà không báo lỗi

Vị trí: `sync_nvl_stock.py:231` và `:307`.

Nguồn tổng hợp có VT001=100 và VT002=200: xóa `<dimension>` gây TypeError vì `max_row=None`; đặt dimension ngắn hơn dữ liệu thật khiến reader chỉ trả VT001, mất VT002. Điều này có thể khiến tồn VT002 ở đích bị giữ cũ dưới dạng cảnh báo missing. Cả reader nguồn và đích đều dùng pattern này.

Sửa: dùng reader streaming theo hàng, xử lý dimension thiếu/sai (`reset_dimensions()` hoặc duyệt XML đến hết dữ liệu). Không dùng `range(max_row)` và gọi `.cell()` lặp lại trên read-only worksheet; pattern đó còn đọc lại XML nhiều lần, không đạt O(n+m). Test cả nguồn/đích, dimension thiếu, ngắn hơn dữ liệu và fixture đủ lớn.

### 3. [P1] Không phát hiện nguồn thay đổi sau khi tải snapshot

Vị trí: `sync_nvl_stock.py:673-677`, trước upload tại `:705`.

Mock nguồn đổi từ 100 sang 999 ngay sau download: hàm vẫn upload D2=100 và trả thành công. ETag đích không bảo vệ revision nguồn. Không có bước đọc lại metadata nguồn sau download hoặc trước publish.

Sửa: kiểm tra metadata trước/sau download và recheck nguồn trước upload; revision đổi thì tải snapshot mới và lập lại patch. Kiểm tra lại kết quả đích sau upload. Test nguồn đổi trong lúc tải/tính, đích đổi thực sự nội dung khi 412 (test hiện tại chỉ thay ETag), và không làm mất chỉnh sửa khác của người dùng. Nêu rõ giới hạn không có giao dịch nguyên tử giữa hai file.

### 4. [P2] Report ghi thành công trước khi upload và không cập nhật khi lỗi

Vị trí: `sync_nvl_stock.py:693`, `:705-723`.

Mock upload trả 403: hàm ném lỗi nhưng artifact JSON vẫn có `mode='publish'`, `status='success'`, thông báo “Đồng bộ thành công”. Người đọc artifact có thể hiểu nhầm đã ghi SharePoint. Lỗi đọc/validation xảy ra sớm cũng chưa tạo error report.

Sửa: tách trạng thái đối soát và trạng thái publish; chỉ ghi `published` sau upload và kiểm tra kết quả, `failed` khi lỗi, `dry_run` khi chỉ đề xuất, `unchanged` khi không cần ghi. Ghi error report cho mọi lỗi với phase/attempt/revision và không để proposal/report từ lần cũ giả làm kết quả hiện tại. Test 403, hết retry 412, validation lỗi và timeout.

### 5. [P2] Xác minh ô giữ nguyên ép text sang float khiến dừng cả lần chạy

Vị trí: `sync_nvl_stock.py:521-532`.

Đích gồm một mã khớp cần cập nhật và một mã thiếu nguồn có D='-': patch đúng, giữ nguyên '-', nhưng verify ném `ValueError: could not convert string to float: '-'`. Quy tắc giữ nguyên mã thiếu không yêu cầu ô cũ phải là số.

Sửa: so sánh giá trị theo kiểu hoặc kiểm tra XML ô giữ nguyên, chỉ dùng tolerance khi cả hai là số hữu hạn. Test text, blank, lỗi Excel ở ô không thuộc patch, đồng thời kiểm tra cả unchanged cells và nội dung khác trong chính sheet Ton_NVL.

### 6. [P2] Retry chỉ bao upload GraphRequestError, bỏ qua lỗi đọc và network

Vị trí: `sync_nvl_stock.py:673-677` nằm ngoài try; `:711` chỉ catch `GraphRequestError`.

Qua đọc code: 429/503 tại metadata/download và `requests.Timeout`/`ConnectionError` sẽ thoát ngay dù `is_retryable_graph_error()` hiện có hỗ trợ network. Trường hợp upload timeout sau khi server đã ghi không được đọc lại để xác định kết quả.

Sửa: bao toàn bộ chu trình snapshot/publish bằng retry có phân loại và giới hạn; giữ nguyên fail-fast với lỗi auth/cấu hình/dữ liệu. Sau upload không rõ kết quả phải đọc lại và đối soát trước lần upload tiếp. Test 429 ở GET, 503 ở download, timeout trước/sau khi server ghi. Đây là phát hiện từ code, chưa nằm trong probe tổng hợp đã chạy.

## Những phần chưa nghiệm thu

- `nvl_stock_config.json` còn `target.sharepoint_path=""`; workflow online chưa chạy được với config hiện tại. CLI chỉ hỗ trợ path; không đưa driveItem ID vào trường path như PR description gợi ý.
- Dòng bắt đầu đang mặc định 2; cần xác minh tiêu đề, dòng dữ liệu, mã/đơn vị và công thức D trên workbook thật.
- Chưa thấy kiểm tra merge/protection hoặc xử lý dimension khi tạo D mới. Cần bổ sung theo plan trước tuyên bố bảo toàn đầy đủ.
- Workflow hiện chỉ chạy thủ công, chưa có schedule hay repository_dispatch. Điều này phù hợp giai đoạn dry-run; không được mô tả là đã tự động đồng bộ.
- `verify_nvl_patched_workbook()` mới so ZIP parts khác sheet, chưa chứng minh mọi nội dung khác bên trong Ton_NVL giữ nguyên. Không nên tuyên bố “100%” từ bộ kiểm tra hiện tại.

## Bước tiếp theo

1. Antigravity sửa các mục trên trên chính nhánh `feat/sync-nvl-stock`, bổ sung regression test và push cập nhật PR #22; cập nhật PR description phản ánh đúng phạm vi đã nghiệm thu. Chưa merge/publish/bật lịch.
2. Chạy lại toàn bộ test và probe. Bàn giao commit mới và kết quả; review lại các lỗi P1/P2 đã sửa.
3. Xác minh đường dẫn file đích và cấu trúc hai workbook qua kết nối có quyền. Hoàn thiện config, không đoán thư mục từ URL Doc.aspx. Nếu chưa kết nối được, dùng bản tải về để chạy offline.
4. Chạy dry-run với workbook thật, lấy proposal + report. Đối chiếu tối thiểu 10 mã, kiểm tra số 0/âm/lẻ nếu có, mã thiếu/trùng; mở proposal bằng Excel, kiểm tra không repair và công thức phụ thuộc D.
5. Khi review đạt và config đúng, merge PR để workflow thủ công có trên nhánh mặc định. Chạy lại dry-run bằng GitHub Actions (`publish=false`) với secrets thật, rồi chạy nghiệm thu publish trên bản sao trước lần ghi file chính thức.
6. Sau publish chính thức, tải lại đích đối chiếu D. Chạy lại không đổi phải không upload. Khi đạt mới chốt giờ và bổ sung schedule/Power Automate.

## Bằng chứng cục bộ

- Test suite: `python -X utf8 -m unittest discover -s tests -v`; log `temp_audit/nvl_review_tests.log`.
- Probe tái hiện mục 1–5: `python -X utf8 temp_audit/nvl_review_probe.py`; kết quả `temp_audit/nvl_review_probe.json`.
- Thư mục `temp_audit` đang được gitignore; không có dữ liệu nghiệp vụ thật trong các probe.

Prompt giao Antigravity:

> Đọc REVIEW_PR22_TON_NVL.md và sửa các phát hiện review trên PR #22, nhánh feat/sync-nvl-stock. Tái hiện lỗi, sửa code, thêm regression tests và chạy toàn bộ suite. Xác minh cấu hình SharePoint khi có quyền, chạy dry-run dữ liệu thật nếu truy cập được. Push cập nhật vào PR hiện tại, báo commit mới, test và phần còn chưa nghiệm thu. Chưa merge, chưa publish workbook chính thức, chưa bật lịch tự động.
