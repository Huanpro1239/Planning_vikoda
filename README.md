# Planning_vikoda

Hệ thống lập kế hoạch sản xuất Vikoda:

**SharePoint → dữ liệu tồn/FC/nợ → Planning Engine → Proposal → Review → Controlled Publish**

## Thuật toán hiện hành

Engine: `ke_hoach_sx_tuan_v2_service_first`

### 1. Service First

Khi capacity không đủ:

1. **FC và Nợ kho được ưu tiên trước.**
2. **Tồn cuối dự kiến là Safety-stock buffer.**
3. Buffer chỉ được sản xuất bằng **capacity còn lại**.
4. Không đủ buffer nhưng vẫn đủ Service → có thể `ready_for_publish`.
5. Không đủ Service → `review_required`.

Rule Nợ kho `SUBTRACT_BOOK_ON_DEBT` và `IGNORE_BOOK_ON_DEBT` vẫn lấy từ `Danh_muc`.

### 2. KHS + PET 9000

KHS và PET 9000 là **một máy vật lý chung**:

- Một timeline chung.
- Không chạy song song.
- Một capacity ca/ngày chung.
- Đổi Quy cách giữa campaign liên tiếp tính setup.
- Service của toàn bộ SKU được xếp trước Safety-stock buffer.

## Ý nghĩa cột kế hoạch

Trong `Ke_hoach_SX`:

- **O**: nhu cầu đầy đủ, gồm mục tiêu tồn cuối dự kiến theo rule.
- **P**: sản lượng sản xuất cam kết thực tế sau khi xét capacity.
- **Q**: số ngày SX tương ứng với P.
- **R**: ngày bắt đầu SX.
- **S:...**: lịch SX ngày.

## Trạng thái

- `ready_for_publish`: Service đủ, resource hợp lệ, metadata hợp lệ.
- `review_required`: thiếu Service, lỗi resource hoặc metadata.
- `safety_stock = partially_achieved`: bán hàng đủ nhưng chưa đạt tồn cuối dự kiến.

## Vận hành

### Tạo Proposal

GitHub Actions → **Sync SharePoint Stock**:

- `publish = false`

Workflow chỉ đọc SharePoint, tính workbook và tạo artifact review.

### Publish

Chỉ sau khi proposal được kiểm tra:

- `publish = true`

Publish dùng **ETag / If-Match**. Nếu dữ liệu thay đổi trong lúc tính, pipeline đọc lại snapshot và tính lại.

## CI

```bash
python -m unittest discover -s tests -v
```

Repository chỉ giữ engine và test của luồng production hiện hành. Các scheduler V2–V10, dry-run và workflow thử nghiệm cũ đã được loại khỏi nhánh hiện hành.
