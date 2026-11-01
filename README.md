# Planning_vikoda – Tổng hợp Sell In, Xuất kho & NXT kế toán

Dựng lại `Planning App/Clean Data/Data ban hang/Tong_Hop_Ban_Hang.xlsx` trên SharePoint Planning
từ toàn bộ file `Planning App/Bao cao ban hang/Sell In` và `Planning App/Danh muc/Danh Muc San Pham.xlsx`.

## Quy tắc

- Công ty đọc từ ô A1 (Tên DN): có "VKD" → VKD, có "KHÁNH HÒA" → Vikoda. Kỳ đọc từ dòng "Từ ngày".
- **VKD**: chỉ lấy `MaSanPhamMoi` bắt đầu bằng `2`.
- **Vikoda**: chỉ lấy `MaSanPhamMoi` bắt đầu bằng `1` và `MaKhachHangMoi` khác `VKD3`.
- Không lấy mã `130100014` / `230100014` (Nước khoáng bán thành phẩm).
- Quy đổi: tổng `SoLuong` theo sản phẩm/tháng ÷ `Quy cách` (Danh mục của đúng công ty) → Thùng / Két / Bình.
- Sheet **Pivot** gộp Vikoda + VKD theo phần đuôi mã (`1xxxxxxxx` = `2xxxxxxxx`), cột theo tháng `Txx/yyyy`, cột cuối `Tong`.
- Sheet **Data**: số liệu gọn tách theo công ty. Sheet **Log**: từng file nguồn và mã thiếu quy cách.

## Cách chạy

| Thành phần | Vai trò |
|---|---|
| Power Automate "Sell In - tổng hợp bán hàng tự động" | Khi có file Sell In mới/sửa → gửi `repository_dispatch` `sharepoint_sell_in_updated` |
| Workflow `Sync Sell In Tong Hop` (`.github/workflows/sync-sell-in.yml`) | Chạy `python -m sell_in`; thêm lịch dự phòng 07:10 sáng VN và chạy tay |
| `sharepoint/` | Đăng nhập Microsoft Graph (GitHub OIDC → Entra) và đọc/ghi SharePoint |
| `sell_in/` | Logic tổng hợp (`core.py`) và orchestration SharePoint (`__main__.py`) |

Secrets cần có: `MS_TENANT_ID`, `MS_CLIENT_ID` (federated credential cho repo).

```bash
make install
make test          # test offline
make sell-in-dry   # tính nhưng không ghi SharePoint
```

## Xuất kho (`xuat_kho/`)

Dựng lại `Planning App/Clean Data/Data Xuat kho/Tong_Hop_Xuat_Kho.xlsx` từ `Planning App/Bao cao xuat kho/<Vikoda|VKD>/*.xlsm`
(sheet `REPORT_XUATBANHANG`), workflow `Sync Xuat Kho Tong Hop` (`.github/workflows/sync-xuat-kho.yml`), kiểm tra mỗi 15 phút.

- Mã lấy giống Sell In: VKD mã `2…`, Vikoda mã `1…` và `MAKHACHHANG` khác `VKD3`; bỏ `130100014` / `230100014`.
- Chỉ lấy **xuất kho thực tế**: bỏ dòng `CHUYENKHOKYGUI` = "Chuyển qua kho ký gửi" (chỉ là chuyển sang kho ký gửi);
  hàng ký gửi được tính khi xuất từ KHO KÝ GỬI (`LOAIDONHANG` = "Ký gửi") ở kỳ thực xuất.
- Quy đổi `SOLUONG ÷ Quy cách`. Pivot gộp 2 công ty, mọi loại đơn; sheet Data tách theo `LoaiDonHang`.

```bash
python -X utf8 -m xuat_kho --dry-run
```

## NXT kế toán + tồn ký gửi (`nxt_ke_toan/`)

Dựng lại `Planning App/Clean Data/Data NXT ke toan/Tong_Hop_NXT_Ke_Toan.xlsx` từ
`Planning App/Bao cao ton kho/Ton ke toan/<Vikoda|VKD>/Thang N.xlsm` + `Bao cao xuat kho` + Danh mục,
workflow `Sync NXT Ke Toan Tong Hop` (`.github/workflows/sync-nxt-ke-toan.yml`), kiểm tra mỗi 15 phút.

- Thành phẩm (Vikoda, mã `1…`) / Hàng hóa (VKD, mã `2…`) quy đổi Thùng/Két/Bình, gộp 2 công ty theo đuôi mã; bỏ `130100014` / `230100014`.
- Kế toán ghi "Chuyển qua kho ký gửi" là xuất → **Tồn thực = Tồn kế toán + Tồn ký gửi**;
  Tồn ký gửi = đầu kỳ + chuyển vào ký gửi − xuất từ KHO KÝ GỬI (từ báo cáo Xuất kho).
  Đầu kỳ ký gửi hiện là *ước tính tối thiểu* (mức nhỏ nhất để không âm).
- Tồn kho thông minh: xuất thực BQ/ngày 3 tháng; nhu cầu = xuất thực VKD; trạng thái theo số ngày tồn thực
  (Thiếu < 7, Thấp 7–15, An toàn 15–30, Cao 30–60, Dư > 60, Không luân chuyển).
- Sheet: Ton thong minh, Pivot Ton thuc, Pivot Ton ky gui, Pivot So ngay ton, Ky gui, Data, Kiem tra, Log, Huong dan.

```bash
python -X utf8 -m nxt_ke_toan --dry-run
```

## FIT – ĐH Bán trong kỳ NPP → SharePoint, điều khiển qua Telegram (`fit_ban_hang/`)

Đăng nhập FIT (`https://103.68.251.239:8888/FIT/`), mở **Báo cáo kế toán › Báo cáo khác › ĐH Bán trong kỳ NPP**,
đặt Từ ngày/Đến ngày, Xem thời gian = *Ngày hóa đơn*, **Không pivot** = bật, Lấy chi nhánh = tắt, các ô khác để mặc định,
rồi bấm **Tạo báo cáo** hai lần và ghi đè 2 file lên SharePoint Planning:

| Mức lấy | File SharePoint |
|---|---|
| Đơn hàng | `BCDonHangBanTrongKyNPP.xlsm` |
| Hóa Đơn | `Lay So Hoa don.xlsm` |

Thư mục mặc định `Planning App/Bao cao ban hang/DH Ban trong ky NPP` (đổi bằng biến `FIT_SP_FOLDER`).
Đuôi file giữ theo file FIT trả về (FIT xuất `.xlsm`); nếu FIT trả `.xlsx/.xls` thì tên giữ nguyên, đuôi theo file thật.

**CAPTCHA:** FIT bắt nhập CAPTCHA mỗi lần đăng nhập, không tự giải. Bot gửi ảnh CAPTCHA vào Telegram, bạn trả lời
ảnh bằng mã (chờ tối đa 5 phút, `/huy` để dừng). Sai mã được nhập lại 1 lần. Xong việc bot **đăng xuất FIT**.
Nếu FIT báo tài khoản đang đăng nhập nơi khác, hãy Đăng xuất FIT trên trình duyệt của bạn rồi gửi lại lệnh.

**Telegram** (chat riêng, chỉ ID trong `TELEGRAM_ALLOWED_USER_IDS`):

```text
/banhang                         đầu tháng -> hôm nay
/banhang thangtruoc              cả tháng trước
/banhang 10/2026                 cả tháng
/banhang 01/10/2026 15/10/2026   khoảng ngày
/huy  /help  /id
```

Lệnh được nghe bởi job `Auto Sync SharePoint (15 phut)` trong thời gian nghỉ giữa các vòng (phản hồi gần như ngay).
Bot gửi lại 2 file vào chat và link SharePoint; lỗi sẽ kèm ảnh màn hình FIT (chỉ khi đã đăng nhập).
**Không chạy bot `erp_bot` trên máy bằng cùng token** — hai nơi cùng đọc tin nhắn sẽ tranh nhau (lỗi 409).

Cấu hình (Settings → Secrets and variables → Actions):

| Loại | Tên | Giá trị |
|---|---|---|
| Secret | `FIT_USERNAME` | tài khoản FIT |
| Secret | `FIT_PASSWORD` | mật khẩu FIT |
| Secret | `TELEGRAM_BOT_TOKEN` | token bot từ @BotFather |
| Secret | `TELEGRAM_ALLOWED_USER_IDS` | ID Telegram, ngăn cách dấu phẩy (gõ `/id` cho bot để xem) |
| Variable | `FIT_AUTO_GIO` | (tùy chọn) `07:30` = mỗi sáng tự gửi CAPTCHA lấy kỳ đầu tháng→hôm nay; bỏ qua nếu hôm nay đã có file |
| Variable | `FIT_SP_FOLDER` | (tùy chọn) thư mục SharePoint đích |
| Variable | `FIT_IGNORE_HTTPS_ERRORS` | (tùy chọn) `true` nếu FIT dùng chứng chỉ tự ký |
| Variable | `FIT_REPORT_URL` | (tùy chọn) URL trực tiếp của form báo cáo, thay cho bấm menu |
| Variable | `FIT_DATE_FORMAT` | (tùy chọn) mặc định `%d/%m/%Y` |

Repo đang **public**: tuyệt đối không ghi tài khoản/mật khẩu/token vào mã nguồn; log chỉ in số dòng/kích thước file.

Chạy thử trên máy (nhập CAPTCHA ở terminal, lưu file tại thư mục hiện tại, không ghi SharePoint):

```bash
FIT_USERNAME=... FIT_PASSWORD=... python -X utf8 -m fit_ban_hang run --ky 10/2026 --dry-run --no-telegram
```

Kiểm thử offline dùng FIT mô phỏng (menu trái, form trong iframe, nút Tạo báo cáo tải `.xlsm`): `make test`.
