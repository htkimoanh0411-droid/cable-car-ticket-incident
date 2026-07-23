# Trạm điều phối sự cố soát vé

Ứng dụng web quản lý quy trình xử lý sự cố hệ thống soát vé: kích hoạt sự cố → xác nhận tham gia → lập biên bản → bàn giao BCNTT → đối soát dữ liệu → hoàn tất & đóng → báo cáo.

## Cách chạy

Chỉ cần **1 file `index.html`** — mở trực tiếp bằng trình duyệt là chạy được, không cần cài đặt gì.

## Deploy lên GitHub Pages

1. Tạo repository mới trên GitHub, đẩy 2 file `index.html` và `README.md` lên.
2. Vào **Settings → Pages**, chọn nhánh (`main`) và thư mục gốc (`/root`), lưu lại.
3. Sau 1–2 phút, truy cập đường link GitHub cấp (dạng `https://<tên-user>.github.io/<tên-repo>/`).

## Deploy lên Vercel

1. Import repository vào Vercel.
2. **Framework Preset**: chọn **Other** (đây là trang tĩnh, không cần build).
3. Để trống lệnh build, **Output Directory** để mặc định (thư mục gốc chứa `index.html`).
4. Deploy — Vercel sẽ tự phục vụ `index.html`.

> Lỗi `404 NOT_FOUND` trên Vercel trước đây là do đẩy file `.jsx` thô (không phải trang HTML). File `index.html` này khắc phục hoàn toàn.

## Lưu trữ dữ liệu

Dữ liệu sự cố được lưu bằng **localStorage của trình duyệt** — tức lưu riêng trên từng máy/trình duyệt truy cập. Ưu điểm: chạy ngay, miễn phí, không cần server. Hạn chế: dữ liệu **không tự đồng bộ giữa các máy khác nhau**.

Nếu sau này cần nhiều người dùng chung một kho dữ liệu thời gian thực (mỗi bộ phận một máy nhưng nhìn chung một bảng sự cố), cần thay lớp lưu trữ bằng một backend (ví dụ Firebase Firestore, Supabase). Điểm cần sửa nằm ở 2 hàm `load()` và `persist()` trong mã nguồn — hiện đang gọi `localStorage.getItem`/`setItem`.

## Thư viện dùng qua CDN

- React 18 (unpkg)
- SheetJS/xlsx (jsdelivr) — đọc/ghi file Excel để đối soát và xuất báo cáo

Cần có kết nối internet để tải các thư viện này khi mở trang.
