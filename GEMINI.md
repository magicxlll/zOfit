# zOfit QS Pro — Quy Tắc Điều Hành AI Agent (GEMINI.md)

Tài liệu này là quy tắc bất biến dành cho mọi AI Agent (Antigravity, Gemini, Claude, Codex, Cursor...) khi làm việc với dự án **zOfit QS Pro**.

---

## 1. QUY TẮC BẮT ĐẦU PHIÊN (Session Initialization & Đồng Bộ Đa Máy)

Khi người dùng yêu cầu: **"Đọc file memory dự án và tiếp tục phát triển dự án"** (hoặc bắt đầu một phiên làm việc mới trên bất kỳ máy tính nào):

1. **Kiểm tra và Cập nhật Git Repository Mới Nhất**:
   - Nếu thư mục hiện tại chưa có kho Git:
     ```powershell
     git clone https://github.com/magicxlll/zOfit.git .
     ```
   - Nếu đã có kho Git:
     ```powershell
     git fetch origin
     git status
     git pull origin main
     ```
     Đảm bảo luôn làm việc trên nhánh `main` và có mã nguồn mới nhất từ `https://github.com/magicxlll/zOfit.git`.

2. **Nạp Ngữ Cảnh Dự Án Theo Thứ Tự Ưu Tiên**:
   - Đọc kỹ `memory.md`:
     - Xem **Mục 2. Trạng thái hiện tại (Snapshot)** để biết phiên bản, kết quả build/test gần nhất.
     - Xem **Mục 5. Bài học kỹ thuật & Bẫy đã xử lý** để tuyệt đối không lặp lại lỗi cũ.
     - Xem **Mục 6. Nhật ký phiên làm việc** để biết phiên trước vừa làm gì.
   - Đọc `tasks/todo.md` để xác định các task đang chờ triển khai.
   - Đọc phần đặc tả liên quan trong `docs/spec.md` nếu động vào nghiệp vụ QS / BIM / CAD / Rebar.

---

## 2. QUY TẮC TỰ ĐỘNG SAO LƯU SAU MỖI PHIÊN DEBUG / PHÁT TRIỂN (Auto-Backup & Sync)

Sau mỗi phiên debug, sửa lỗi, hoàn thành một tính năng, hoặc trước khi kết thúc phiên:

1. **Xác Minh Chất Lượng (Verification Gate)**:
   - Chạy kiểm thử tự động của Engine:
     ```powershell
     dotnet test engine/zOfit.Engine.Tests/zOfit.Engine.Tests.csproj
     ```
   - Nếu có thay đổi ở mã nguồn VBA hoặc Engine cần phát hành: Chạy `build/build_release.ps1` để kiểm tra E2E COM.
   - Mọi unit test phải **PASS 100% (0 errors)** trước khi commit.

2. **Cập Nhật Bộ Nhớ Dự Án**:
   - Cập nhật `memory.md`: Cập nhật ngày giờ, trạng thái snapshot ở Mục 2, và ghi lại phiên làm việc ở Mục 6.
   - Cập nhật tiến độ `[x]` trong `tasks/todo.md`.

3. **Tự Động Commit & Push Lên GitHub**:
   - Thực hiện lệnh:
     ```powershell
     git status
     git add .
     git commit -m "<type>(<scope>): <tóm tắt ngắn gọn thay đổi>"
     git push origin main
     ```
   - Báo cáo rõ ràng commit hash và link repo cho người dùng.

---

## 3. CÁC NGUYÊN TẮC KỸ THUẬT BẤT DI BẤT DỊCH (Technical Invariants)

1. **100% Công Thức Excel Sống (Pure Excel Formula Output)**:
   - Toàn bộ bảng tính chi tiết và bảng thống kê thép phải dùng 100% công thức Excel sống tương thích Excel 2007–365.
   - Tuyệt đối không dùng hàm mới (`LET`, `LAMBDA`, `XLOOKUP`, `IFS`) hoặc VBA UDF trong ô kết quả gửi cho CĐT/Sếp.
   - Tuyệt đối không xuất "số chết".

2. **Mã Nguồn VBA Thuần ASCII (Pure ASCII)**:
   - Tất cả các file `.bas`, `.cls`, `.txt` trong `src/vba/` chỉ chứa ký tự mã ASCII $\le 127$.
   - Chuỗi tiếng Việt có dấu phải mã hóa qua hàm `DecodeUtf8(Chr(...))` hoặc Unicode Escape.

3. **AutoCAD MCP Live Integration**:
   - Toàn bộ logic quét bảng CAD, nhận diện tọa độ, dịch mã TCVN3 sang Unicode, bóc tách tiết diện và bố trí thép được xử lý 100% bằng code C# .NET 8 cục bộ qua cổng TCP 9877.
   - Hoàn toàn độc lập, không yêu cầu kết nối mạng, không gửi dữ liệu ra ngoài, và không phụ thuộc dịch vụ AI đám mây.

4. **Bảo Vệ Kích Thước File Git**:
   - Tuyệt đối không commit file nhị phân lớn $> 100\text{ MB}$ (file bộ cài `dist/setup/zOfit_QS_Pro_Setup.exe`, `zofit-engine.exe`, thư mục `bin/`, `obj/`, `.vs/`).
   - Luôn duy trì file `.gitignore` chuẩn hóa.
