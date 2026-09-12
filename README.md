# Web Quản lý Hóa đơn (Đầu vào / Đầu ra) + Tạo hóa đơn bán hàng

## Chức năng
1. **Bảng kê thuế GTGT** — nhập từng hóa đơn đầu vào/đầu ra, xuất Excel theo mẫu
   bảng kê (giống Phụ lục 01-1/GTGT), có sheet đối chiếu tự tính chênh lệch.
2. **Tạo hóa đơn bán hàng** — nhập nhiều mặt hàng, xuất PDF in được (giống mẫu
   "Hóa đơn bán hàng" viết tay/tự in mà bạn hay gặp ở các hộ kinh doanh), tự đọc
   số tiền bằng chữ.
3. **Phân quyền**: Admin (toàn quyền, quản lý tài khoản) / Nhân viên (nhập liệu).

## Chạy thử ở máy local
```bash
pip install -r requirements.txt
python3 app.py
```
Mở trình duyệt: http://127.0.0.1:5000
Tài khoản mặc định: **admin / admin123** — đổi mật khẩu ngay bằng cách vào mục
"Tài khoản" → xóa/tạo lại, hoặc tự sửa trong DB.

## Deploy lên Railway
1. Đẩy toàn bộ thư mục này lên một GitHub repo mới.
2. Vào railway.app → New Project → Deploy from GitHub repo → chọn repo vừa tạo.
3. Railway tự nhận `requirements.txt` và `Procfile` để chạy bằng gunicorn.
4. Vào tab **Variables**, thêm biến `SECRET_KEY` = một chuỗi ngẫu nhiên bất kỳ
   (bảo mật session đăng nhập).
5. Railway sẽ build và cấp domain dạng `xxxx.up.railway.app`.

### Lưu ý về dữ liệu (SQLite)
- File `hoadon.db` được tạo tự động trong container khi chạy lần đầu.
- **Quan trọng**: Railway free/hobby plan không đảm bảo ổ đĩa persistent theo
  mặc định — nếu container bị redeploy, dữ liệu SQLite có thể mất. Để an toàn:
  - Vào Railway → thêm một **Volume** (Settings → Volumes) và mount vào đường
    dẫn chứa `hoadon.db` (ví dụ `/app/data`), rồi sửa `DB_PATH` trong `app.py`
    trỏ vào thư mục đó.
  - Hoặc định kỳ tải file `hoadon.db` về backup (giống cách bạn đang làm với
    SGV/PythonAnywhere).

## Cấu trúc thư mục
```
hoadon_web/
├── app.py                  # toàn bộ logic Flask (routes, DB, export)
├── vn_number_to_words.py   # đọc số tiền thành chữ tiếng Việt
├── requirements.txt
├── Procfile                 # lệnh chạy cho Railway (gunicorn)
├── templates/                # giao diện HTML (Jinja2 + Bootstrap 5)
└── static/fonts/              # font DejaVu Sans hỗ trợ tiếng Việt cho PDF
```

## Nâng cấp có thể làm thêm sau
- Import hóa đơn từ ảnh chụp (OCR) thẳng vào bảng kê hoặc form tạo hóa đơn.
- Cho phép upload logo cửa hàng để in trên PDF.
- Xuất file Excel/PDF theo quý/năm thay vì chỉ theo tháng.
- Thêm log lịch sử sửa/xóa hóa đơn (audit trail) để phục vụ kiểm tra nội bộ.
