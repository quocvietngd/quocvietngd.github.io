- [x] Verify that the copilot-instructions.md file in the .github directory is created.

- [x] Clarify Project Requirements
	- Vite + Vanilla JavaScript dashboard.
	- Bao gồm phân quyền, data source remote, xuất PDF.

- [x] Scaffold the Project
	- Đã scaffold thủ công tại thư mục hiện tại do thiếu `npx`.
	- Đã tạo `index.html`, `src/main.js`, `src/style.css`, `package.json`, `README.md`.

- [x] Customize the Project
	- Đã thêm phân quyền role (CEO/Trưởng bộ phận).
	- Đã thêm đồng bộ dữ liệu Local/REST API/Google Sheets endpoint.
	- Đã thêm xuất báo cáo PDF bằng `html2canvas` + `jsPDF`.

- [x] Install Required Extensions
	- No extensions needed.

- [x] Compile the Project
	- Đã cài runtime Node portable trong user directory.
	- Đã chạy thành công `npm install` và `npm run build`.


- [x] Create and Run Task
	- Đã tạo task `Run NORA Dashboard` trong `.vscode/tasks.json`.
	- Đã cập nhật task để nạp `PATH` runtime portable trước khi chạy.

- [x] Launch the Project
	- Đã launch thành công bằng task `Run NORA Dashboard` tại `http://localhost:5173/` (non-debug mode).

- [x] Ensure Documentation is Complete
	- Đã có `README.md` và `.github/copilot-instructions.md` cập nhật theo trạng thái hiện tại.
	- Đã xóa toàn bộ HTML comments khỏi file này.
