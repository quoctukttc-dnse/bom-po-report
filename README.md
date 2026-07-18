# Dashboard tiến độ tạo BOM & PO

Dashboard tĩnh (chạy hoàn toàn trên trình duyệt) để theo dõi tiến độ tạo BOM và PO trên hệ thống: lọc theo thời gian, người tạo, khách hàng, loại BOM/PO, stage, trạng thái; xem theo ngày/tuần/tháng và theo khung giờ upload SAP/Centric.

## Cấu trúc

| File | Vai trò |
|---|---|
| `index.html` | Toàn bộ dashboard (HTML + CSS + JS) |
| `data.js` | Dữ liệu mặc định được nhúng sẵn (sinh từ file Excel) |
| `tools/update_data.py` | Script tạo lại `data.js` từ file `BOM PO Status.xlsx` mới |

## Đưa lên GitHub Pages (làm 1 lần)

1. Đăng nhập GitHub → bấm **New repository** → đặt tên ví dụ `bom-po-dashboard`.
   - Chọn **Private** nếu không muốn ai ngoài bạn xem được code/dữ liệu (GitHub Pages của repo private cần gói trả phí; repo **Public** thì Pages miễn phí — lưu ý dữ liệu trong `data.js` sẽ công khai).
2. Trong repo mới → **Add file → Upload files** → kéo toàn bộ file trong thư mục này lên (`index.html`, `data.js`, thư mục `tools/`, `README.md`) → **Commit changes**.
3. Vào **Settings → Pages** → mục *Build and deployment*:
   - Source: **Deploy from a branch**
   - Branch: **main** / thư mục **/(root)** → **Save**
4. Chờ 1–2 phút, dashboard sẽ có tại:
   `https://<tên-user>.github.io/bom-po-dashboard/`

## Cập nhật dữ liệu

Có 2 cách, dùng cách nào cũng được:

**Cách 1 — Kéo-thả (nhanh nhất, không cần cài gì):** mở dashboard, kéo file `BOM PO Status.xlsx` mới xuất từ hệ thống thả vào ô ở góc phải trên. Dữ liệu cập nhật ngay trên trình duyệt (chỉ trong phiên xem đó, không ghi vào repo).

**Cách 2 — Cập nhật dữ liệu mặc định trong repo:** chạy

```bash
pip install pandas openpyxl
python3 tools/update_data.py "đường/dẫn/BOM PO Status.xlsx"
```

rồi commit + push file `data.js` mới (hoặc upload đè `data.js` qua giao diện web GitHub). Mọi người mở link sẽ thấy dữ liệu mới mà không cần thao tác gì.

## Ghi chú về dữ liệu

- Cột **Created Date** trong file xuất chỉ có ngày, không có giờ, nên biểu đồ "theo khung giờ" dựa trên timestamp **Upload SAP Date / Upload Centric Date**. Nếu sau này file xuất có kèm giờ ở Created Date, dashboard/script vẫn đọc bình thường.
- File xuất có lỗi ngày tháng bị đảo (ngày ≤ 12 bị Excel hiểu nhầm định dạng Mỹ mm/dd) — cả `tools/update_data.py` lẫn phần kéo-thả đều tự sửa lại.
- Các dòng ghi chú (remark) tràn dòng trong sheet BOM được tự động loại bỏ.
- `data.js` chỉ chứa các trường phục vụ biểu đồ tổng hợp (ngày tạo, người tạo, khách hàng, loại, stage/trạng thái, giờ upload) — **không** chứa mã sản phẩm, số BOM/PO, nhà cung cấp hay giá trị đơn hàng, nên an toàn khi đặt trên repo public.
- Dữ liệu chỉ được xử lý trong trình duyệt của người xem, không gửi đi máy chủ nào.
