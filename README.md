# Sổ đo thủy chuẩn hạng 4

## Dùng nhanh

Mở file **`so_do.html`** bằng trình duyệt (Chrome, Edge, Firefox). Không cần cài Python hay internet.

1. Nhập thông tin dự án và số tổng hợp
2. Xem preview phân bổ từng trạm và kiểm tra hạng IV
3. Bấm **Tải file Excel** → `So_do_(tên_mốc_mua).xlsx`

Phần mềm **tự phân bổ** số đo sao cho mọi trạm đạt `|H−mid stadia| ≤ 5 mm`, `|J| ≤ 2 mm`, đồng thời khớp tổng L/H/I bạn nhập.

## Tái tạo file HTML (khi sửa giao diện hoặc logic)

```bash
python build_html.py
```

Cần có `vendor/exceljs.min.js` (script build sẽ báo nếu thiếu).

File **`So do.xlsx`** là sổ mẫu tham khảo — không được nhúng vào chương trình HTML.
