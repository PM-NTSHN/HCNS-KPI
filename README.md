# HCNS-KPI · Tool tổng hợp Báo cáo KPI tháng

Tool HTML chạy offline (một file duy nhất) giúp bộ phận Nhân sự – Công ty Nam Trường Sơn Hà Nội tổng hợp báo cáo KPI tháng từ các file Excel xuất từ Base Wework.

## Dùng ngay

- **Online:** mở trang GitHub Pages của repo (index.html).
- **Offline:** tải file `Tool tổng hợp KPI.html` (hoặc `index.html`), mở bằng Chrome/Edge. Không cần cài đặt, không cần internet; dữ liệu chỉ xử lý trên máy người dùng.

## Chức năng

1. Nạp tối đa 20 file Excel raw (.xlsx/.xls): file công việc từng phòng ban (VD `T9 admin.xlsx`, `T9 Tech 1.xlsx`) và file dự án (VD `T9 dự án.xls`).
2. Tự tách cột "Trường dữ liệu tùy chỉnh" của file dự án thành Mức độ / Đánh giá chất lượng…, gộp vào sheet `Data` theo phòng ban của từng nhân sự.
3. Xuất file `Báo cáo KPI Tháng <tháng.năm>.xlsx` theo form mẫu:
   - `Mastersheet`: dashboard KPI, bảng & biểu đồ theo phòng ban, tra cứu theo nhân sự, Top 5 / Top 3, insight tự động, lọc theo giá trị "Đánh giá chất lượng".
   - `Nguyên tắc đánh giá`, một sheet cho mỗi phòng ban, `Tổng hợp Dự án`, `Data`, `Data_Dự án`.
   - Số liệu là công thức Excel, tự cập nhật khi sửa sheet Data.
4. Xem trước trên trình duyệt: tổng quan, tra cứu nhân sự, lọc đánh giá chất lượng, Top & Insight, kiểm tra chất lượng dữ liệu.

## Cấu trúc mã nguồn

| File | Nội dung |
|---|---|
| `src/tool_src.html` | Giao diện (CSS, khung HTML) |
| `src/core.js` | Đọc file, chuẩn hoá, tính toán, insight, cảnh báo |
| `src/excel.js` | Dựng các sheet Excel theo form mẫu và Mastersheet |
| `src/charts.js` | Chèn biểu đồ Excel (DrawingML) |
| `src/app.js` | Xử lý giao diện |
| `src/libs/` | SheetJS, ExcelJS, JSZip (nhúng inline khi build) |
| `src/assets/` | Logo NTS |
| `src/build.py` | Ghép thành `Tool tổng hợp KPI.html` và `index.html` |

Sửa mã trong `src/` rồi build lại:

```bash
py src/build.py
```

## Bảo mật dữ liệu

Repo **không** chứa dữ liệu nhân sự: file raw, file báo cáo `.xlsx` và tài liệu `.docx` đã được loại bằng `.gitignore`. Không commit các file này lên repo public.
