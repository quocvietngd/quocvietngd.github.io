# NORA KPI Dashboard

Dashboard web theo dõi chỉ số phòng ban cập nhật hằng ngày.

## Tính năng chính

- Phân quyền đăng nhập demo theo vai trò: CEO, Trưởng bộ phận.
- Cập nhật chỉ số khi nộp báo cáo hằng ngày.
- Kết nối dữ liệu thật theo 2 lựa chọn:
  - Backend REST API (GET/POST JSON).
  - Google Sheets Web App (Apps Script endpoint).
- Xuất toàn bộ dashboard thành PDF.

## Chạy dự án

1. Cài Node.js LTS.
2. Cài dependencies:
   - `npm install`
3. Chạy môi trường phát triển:
   - `npm run dev`
4. Chạy Telegram webhook bridge (cho realtime Telegram):
  - `npm run dev:webhook`
5. Hoặc chạy cả dashboard + webhook cùng lúc (macOS/Linux):
  - `npm run dev:realtime`
6. Build production:
   - `npm run build`

## Xuất bản web và chỉnh sửa trực tuyến

Project đã được cấu hình auto deploy lên GitHub Pages qua workflow:
- [.github/workflows/deploy-pages.yml](.github/workflows/deploy-pages.yml)

Các bước bật public site:

1. Đẩy project lên GitHub repository (branch `main`).
2. Vào GitHub repository > Settings > Pages.
3. Ở Source, chọn `GitHub Actions`.
4. Mỗi lần push lên `main`, site sẽ tự build và deploy.
5. Có thể chỉnh sửa trực tuyến ngay trên GitHub (phím `.` hoặc nút Edit file), commit xong là web tự cập nhật.

Ghi chú:
- URL site có dạng `https://<github-username>.github.io/<repo-name>/`.
- Build đã dùng `base: "./"` trong [vite.config.js](vite.config.js), nên hoạt động tốt trên GitHub Pages project URL.

## Telegram Realtime (Webhook)

1. Vào trang Nguồn dữ liệu, phần Đồng bộ từ Telegram.
2. Nhập Bot Token + Chat ID.
3. Nhập Webhook Public URL (bắt buộc HTTPS và Telegram truy cập được).
4. Bấm Kết nối realtime để đăng ký webhook.
5. Sau khi kết nối, hệ thống tự động kéo dữ liệu mới định kỳ, không cần bấm đồng bộ mỗi lần.

Ghi chú:
- Nếu chạy local, cần tunnel HTTPS (ngrok/cloudflared) trỏ về `http://localhost:8787` để Telegram gọi webhook.
- Nút Đọc tin báo cáo vẫn có thể dùng để đồng bộ thủ công ngay lập tức.

## Dữ liệu remote kỳ vọng

GET endpoint trả về mảng JSON:

```json
[
  {
    "date": "2026-04-11",
    "department": "Kỹ thuật",
    "completion": 88,
    "quality": 91,
    "issues": 2,
    "submitter": "Nguyễn Minh",
    "updatedAt": 1770000000000
  }
]
```

POST endpoint nhận 1 object report có cấu trúc tương tự.

## Tài khoản demo

- CEO: passcode `NORA-CEO-2026`
- Trưởng bộ phận: passcode `NORA-HEAD-2026`

Bạn nên thay passcode hoặc tích hợp auth thật qua backend trước khi đưa vào production.
