---
description: Auto Git Sync & Memory Protocol for zOfit QS Pro
always_on: true
---

# Git Sync & Continuity Rules for zOfit QS Pro

## Khi Bắt Đầu Phiên Làm Việc (Mọi Máy Tính)
1. Khi người dùng yêu cầu tiếp tục dự án hoặc đọc `memory.md`:
   - Nếu chưa clone: `git clone https://github.com/magicxlll/zOfit.git .`
   - Nếu đã có thư mục git: `git pull origin main` để kéo code mới nhất.
   - Đọc `memory.md` để nắm Snapshot hiện tại, bài học kỹ thuật và nhật ký phiên trước.

## Tự Động Sao Lưu Sau Mỗi Phiên Debug / Phát Triển
1. Chạy unit tests: `dotnet test engine/zOfit.Engine.Tests/zOfit.Engine.Tests.csproj`.
2. Cập nhật `memory.md` (mục 2 & mục 6) và `tasks/todo.md`.
3. Tự động commit và push:
   ```powershell
   git status
   git add .
   git commit -m "<type>(<scope>): <mô tả>"
   git push origin main
   ```
4. Báo cáo trạng thái commit/push cho người dùng.

## Lưu Ý Về Git & File Lớn
- Không commit các file build nặng > 100MB (`dist/setup/zOfit_QS_Pro_Setup.exe`, `zofit-engine.exe`, `bin/`, `obj/`).
- Duy trì VBA Pure-ASCII và 100% công thức Excel sống.
