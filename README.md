# Planning_vikoda – Tổng hợp Sell In & Xuất kho

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
