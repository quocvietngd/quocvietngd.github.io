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

Chinh sach repo (single source of truth):
- Chi su dung 1 repo duy nhat: `quocvietngd/web-quan-tri`.
- Khong push/deploy sang repo thu 2 de tranh lech phien ban.
- Moi thay doi code va du lieu trien khai deu di qua branch `main` cua repo nay.
- URL public chinh thuc: `https://quocvietngd.github.io/web-quan-tri/`.
- Backend realtime chinh thuc: `https://nora-sync-quocvietngd-2026-2.onrender.com`.

Các bước bật public site:

1. Đẩy project lên GitHub repository (branch `main`).
2. Vào GitHub repository > Settings > Pages.
3. Ở Source, chọn `GitHub Actions`.
4. Mỗi lần push lên `main`, site sẽ tự build và deploy.
5. Có thể chỉnh sửa trực tuyến ngay trên GitHub (phím `.` hoặc nút Edit file), commit xong là web tự cập nhật.

Ghi chú:
- URL site chính thức của project: `https://quocvietngd.github.io/web-quan-tri/`.
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

## Dong bo mat khau da thiet bi (Users Cloud)

Server `server/webhook-server.mjs` da duoc mo rong thanh API luu users + reports dung chung cho nhieu thiet bi.

### 1) Chay API server

- Lenh:
  - `npm run dev:webhook`
- Server se mo cong `8787` va co cac endpoint:
  - `GET/PUT/POST /api/users`
  - `GET/PUT/POST /api/reports`

### 2) Public API de thiet bi khac truy cap

- Dung localtunnel/ngrok/cloudflared tro vao cong `8787`.
- Vi du URL public: `https://abcxyz.loca.lt`

### 3) Cau hinh trong dashboard (tren MOI thiet bi)

- Vao trang `Nguon du lieu`.
- Chon `Data source = REST API`.
- Nhap URL: `https://abcxyz.loca.lt/api/reports`
- Bam `Dong bo du lieu`.

Luu y:
- App tu dong doi `/api/reports` -> `/api/users` de dong bo tai khoan.
- Khi tao/sua tai khoan (doi mat khau), app se day danh sach users len cloud endpoint.
- Khi dang nhap, app se keo users cloud truoc khi kiem tra mat khau.

### 4) Kiem tra nhanh

- `curl https://abcxyz.loca.lt/api/users`
- Neu thay danh sach users JSON va sau khi doi mat khau co cap nhat, dong bo da thanh cong.

## Thiet lap dong bo "nhu phan mem chuan" (khuyen nghi)

Muc tieu: doi mat khau o 1 trinh duyet, tat ca trinh duyet/thiet bi khac tu dong nhan thay doi.

### 1) Deploy backend users sync len cloud (luon online)

Repo da co san file `render.yaml` de deploy len Render.

One-click deploy:
- `https://render.com/deploy?repo=https://github.com/quocvietngd/web-quan-tri`

Buoc nhanh:
- Tao Web Service moi tren Render tu repo nay.
- Render se doc `render.yaml` va chay `node server/webhook-server.mjs`.
- Service co Persistent Disk (`/var/data`) de luu users/reports khong mat sau restart.

Sau khi deploy xong, ban se co URL dang:
- `https://nora-sync-api.onrender.com`

Kiem tra:
- `https://nora-sync-api.onrender.com/api/users`
- `https://nora-sync-api.onrender.com/api/reports`

### 2) Cai endpoint trung tam cho TOAN BO trinh duyet

Sua file `public/runtime-config.json`:

```json
{
  "usersSyncEndpoint": "https://nora-sync-api.onrender.com/api/users"
}
```

Deploy lai frontend (push len `main`).

Loi ich:
- Moi thiet bi mo web deu tu dong nap endpoint nay.
- Khong can nhap lai thu cong tren tung may.

### 3) Hanh vi dong bo sau khi cai xong

- Khi doi mat khau o trang Nhan su: app day users len cloud.
- Cac tab/thiet bi khac tu dong keo users cloud theo chu ky va khi focus lai tab.
- Dang nhap moi luon uu tien users cloud truoc local.

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
