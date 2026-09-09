# Planning_vikoda

Hệ thống lập kế hoạch sản xuất Vikoda:

**SharePoint → dữ liệu tồn/FC/nợ → Planning Engine → Proposal → Review → Controlled Publish**

## Thuật toán hiện hành

Engine: `ke_hoach_sx_tuan_v2_service_first_20260908`

### 1. Service First & Nhu cầu bắt buộc

Khi capacity không đủ:

1. **FC và Nợ kho là nhu cầu dịch vụ bắt buộc (Service Demand).**
2. **Khi FC = 0 nhưng có nợ kho:** Ngày bắt đầu sản xuất mặc định từ ngày đầu tiên của kỳ (ngày 01) thay vì phụ thuộc phép chia tồn/FC, đảm bảo không bỏ sót lịch trả nợ.
3. **Tồn cuối dự kiến là Safety-stock buffer:** Chỉ được sản xuất bằng **capacity còn lại**.
4. **Bảo vệ Service trên timeline:** Buffer của SKU bắt đầu sớm không bao giờ được chiếm lịch hoặc làm thiếu hụt Service của SKU sau trên timeline sản xuất thực tế.
5. **Không đủ buffer nhưng vẫn đủ Service:** `safety_stock = partially_achieved`, có thể `ready_for_publish`.
6. **Không đủ Service (Service Shortfall):** Bắt buộc đánh dấu `service.ok = False`, lưu `stockout_skus`, chuyển trạng thái thành `review_required` và **chặn tuyệt đối auto-publish**.

Rule Nợ kho `SUBTRACT_BOOK_ON_DEBT` và `IGNORE_BOOK_ON_DEBT` được cấu hình từ sheet `Danh_muc`.

### 2. KHS + PET 9000

KHS và PET 9000 là **một máy vật lý chung**:

- Một timeline chung.
- Không chạy song song.
- Một capacity ca/ngày chung.
- Đổi Quy cách giữa campaign liên tiếp tính setup (mặc định 0.5 ca theo PlannerPolicy).
- Service của toàn bộ SKU được ưu tiên 100% trước Safety-stock buffer; công suất còn dư sẽ được phân bổ một phần cho buffer theo thứ tự ưu tiên của SKU mà không chia nhỏ campaign.

## Ý nghĩa cột kế hoạch

Trong `Ke_hoach_SX`:

- **A**: Mã sản phẩm.
- **I**: Số ca theo ngày (shifts_per_day - đầu vào người dùng chỉnh).
- **J**: Tồn đầu thực tế (đồng bộ từ `Ton_kho!D`).
- **K**: Tồn đầu sổ sách (đồng bộ từ `Ton_kho!E:H`).
- **L**: FC tháng được chọn (đồng bộ từ sheet `FC`).
- **M**: Tồn cuối dự kiến (Safety stock target - do pipeline tính toán từ FC, Leadtime và tồn đầu consignment).
- **N**: Nợ kho (đồng bộ từ sheet `No kho!D`).
- **O**: Nhu cầu đầy đủ, gồm mục tiêu tồn cuối dự kiến theo rule.
- **P**: Sản lượng sản xuất cam kết thực tế sau khi xét capacity.
- **Q**: Số ngày SX tương ứng với P.
- **R**: Ngày bắt đầu SX.
- **S:...**: Lịch SX chi tiết theo từng ngày trong tháng.

## Cơ chế phát hiện thay đổi (Change Detection & Loop Prevention)

Hệ thống theo dõi toàn diện các đầu vào nghiệp vụ qua SHA-256 fingerprint:
1. **5 file tồn kho nguồn:** ETag của `actual_stock`, `factory_vikoda`, `factory_vkd`, `accounting_vikoda`, `accounting_vkd`.
2. **Sheet `Danh_muc`:** Toàn bộ cột chính sách (Quy cách/mold, Leadtime, sản lượng/mẻ, sản lượng/ca, Chuyền, Nhóm SP, Phân loại SP, Debt mode / Cách tính nợ, Schedule profile / Profile lịch).
3. **Sheet `FC`:** Selector tháng và bảng số liệu forecast.
4. **Sheet `No kho`:** Danh sách mã và số lượng nợ kho (cột D).
5. **Sheet `Ke_hoach_SX`:** Tập SKU (cột A) và Số ca/ngày (cột I).

> [!NOTE]
> Các cột kết quả do pipeline tính và ghi trên `Ke_hoach_SX` (M, O, P, Q, R, S+) được loại trừ khỏi hash đầu vào nhằm chống hiện tượng kích hoạt lặp (Loop Prevention) sau khi publish.

## Trạng thái và Chính sách Duyệt Publish

- `ready_for_publish`: Service đầy đủ, tài nguyên KHS/PET hợp lệ, metadata hợp lệ. Cho phép tự động publish theo lịch hoặc sự kiện.
- `review_required`: Thiếu Service, hoặc lỗi cấu hình/tài nguyên. **Auto-publish bị chặn hoàn toàn**.
- **Quy trình Duyệt (Review Approval):**
  - Chỉ áp dụng cho trường hợp thiếu Service do giới hạn công suất thực tế (`stockout_risk`).
  - Bắt buộc phải cung cấp đúng `proposal_id` của bản snapshot hiện tại và kèm theo `approval_reason` cụ thể.
  - Lỗi vi phạm máy chung hoặc carryover chưa giải quyết tuyệt đối không thể duyệt bỏ qua.

## Vận hành CI/CD

### 1. Kích hoạt tự động
- **Lịch hàng ngày:** Chạy định kỳ lúc 06:00 sáng VN (23:00 UTC) với cờ `--force`.
- **Power Automate Dispatch:** Sự kiện `repository_dispatch` (`sharepoint_stock_updated`) sử dụng `--skip-if-unchanged` để bỏ qua tính lại nếu không có thay đổi đầu vào.

### 2. Kích hoạt thủ công (Workflow Dispatch)
- **Tạo Proposal (Read-only):** `publish = false`. Chỉ sinh file excel và report artifact trên GitHub Actions để kiểm tra.
- **Duyệt Publish:** `publish = true`, kèm `approval_proposal_id` và `approval_reason` nếu proposal ở trạng thái `review_required`.

### 3. CI Gate
Tất cả các workflow (gồm cả workflow sync và kiểm thử PR) đều chạy qua cổng kiểm thử bắt buộc trước khi chạm vào SharePoint:
```bash
python -X utf8 -m unittest discover -s tests -v
```

## Chạy offline (không cần SharePoint)

Dùng để nghiệm thu/kiểm tra proposal từ các bản sao tải về máy. Script chỉ ĐỌC file nguồn và GHI proposal + report vào thư mục `--out`, không upload, không đổi state gốc:

```bash
python run_offline.py \
    --target "Sắp kế hoạch.xlsx" \
    --actual "Bao cao ton thuc te hien tai.xlsx" \
    --factory-vikoda NXT_Vikoda.xlsm \
    --factory-vkd NXT_VKD.xlsm \
    --accounting-vikoda XNT_ketoan_Vikoda.xlsm \
    --accounting-vkd XNT_ketoan_VKD.xlsm \
    --out ./out --verify
```

## Đồng bộ Tồn Nguyên Vật Liệu (NVL)

Module độc lập `sync_nvl_stock.py` đồng bộ số lượng tồn kho nguyên vật liệu từ `XNT_ketoan_Vikoda.xlsm` (Sheet1, cột B là mã, cột M là tồn) sang `Kế hoạch mua hàng.xlsx` (sheet `Ton_NVL`, ghép mã tại cột A, ghi tồn vào cột D).

### 1. Vận hành Offline
Dùng khi nghiệm thu dữ liệu từ các file tải về máy:
```bash
python -X utf8 sync_nvl_stock.py \
    --config nvl_stock_config.json \
    --source-file "XNT_ketoan_Vikoda.xlsm" \
    --target-file "Kế hoạch mua hàng.xlsx" \
    --out offline_out/nvl
```

### 2. Vận hành GitHub Actions
Workflow riêng `.github/workflows/sync-nvl-stock.yml`:
- **Tạo Proposal (Mặc định):** `publish = false`. Tải snapshot từ SharePoint, đối soát và sinh artifact `nvl_stock_proposal.xlsx` + `nvl_stock_report.json`.
- **Duyệt Publish:** `publish = true`. Tải lên file đích trên SharePoint sau khi proposal đã được kiểm tra.

