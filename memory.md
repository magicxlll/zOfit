# zOfit QS Pro — Project Long-Term Memory (`memory.md`)

> **Mục đích file này**: Bộ nhớ cục bộ lâu dài của dự án **zOfit QS Pro**.
> **Quy tắc bắt buộc cho mọi AI Agent**:
> 1. **Đầu mỗi phiên**: Đọc ngay `memory.md` → `tasks/todo.md` → phần liên quan trong `docs/spec.md` & `tasks/plan.md`.
> 2. **Cuối mỗi phiên (hoặc sau mỗi Checkpoint)**: Bắt buộc cập nhật mục **2. Trạng thái hiện tại**, **6. Nhật ký phiên làm việc**, và đánh dấu tiến độ trong `tasks/todo.md`.

---

## 1. Bản sắc & Mục tiêu cốt lõi của dự án

- **Tên sản phẩm**: **zOfit QS Pro**
- **Mục tiêu cuối cùng**: Hệ thống bóc tách khối lượng xây dựng (QS) bán tự động hóa tối đa cho dự án công nghiệp FDI / dân dụng / hạ tầng, giải quyết bài toán **Single-Entry (Nhập 1 lần → Tự ra toàn bộ bản tính chi tiết & Link BOQ)**.
- **Luồng nghiệp vụ chuẩn**:
  `AutoCAD (Linkcad / Raw Dump)` → `Sheet Data (CAD thô)` → `QS_Data (Nhập 1 lần theo Nhà/Khu vực)` + `QS_Spec (Tiêu chuẩn dự án)` → **1-click sinh thuần công thức Excel sống**:
  1. `<Building>_Detail` (Đào/Lấp đất, Cọc, Bê tông lót, Bê tông kết cấu, Cốp pha — đúng format CĐT `NO|L|W|H|a|b|c|QTY`)
  2. `TKT_<Building>` (Thống kê cốt thép chuẩn 41 hình dạng, link trực tiếp kích thước từ `QS_Data` & quy chuẩn từ `QS_Spec`)
  3. Hoàn thiện (`QS_DoorWin` + `QS_Rooms` → `Finishing_Detail`) & Hạ tầng (`QS_Infra` ← `QS_CAD_Parsed` → `Civil_Detail`)
  4. Map-Link sang BOQ Tổng (`2. ARC+STRUC.`, `3. CIVIL& LANDSCAPE`) giữ nguyên format BOQ CĐT.
- **4 Nguyên tắc bất di bất dịch (User Confirmed)**:
  1. **Chỉ làm việc trong Microsoft Excel** (Tab Ribbon `zOfit QS Pro`) — tuyệt đối không tạo phần mềm tách biệt.
  2. **Kiến trúc Lai + Cài đặt 1-click `.exe`**: Add-in `.xlam` mỏng làm UI + Engine biên dịch `.NET 8` (`zofit-engine.exe`) + Bộ cài `zOfit_QS_Pro_Setup.exe` (người dùng cuối không cần cài Python hay .NET Runtime).
  3. **Đầu ra 100% công thức Excel sống (`.xlsx`, chuẩn Excel 2007–365)**: Sếp/CĐT mở kiểm tra được mọi công thức mà không cần cài Add-in (cấm `LET`/`IFS`/`XLOOKUP`/UDF).
  4. **Thực dụng, không "ảo hóa"**: Ranh giới Tự động / Bán tự động / Thủ công rõ ràng; giữ ô kích thước mở tương thích `Linkcad`; đồng bộ `Sync` không phá hủy dòng sửa tay.

---

## 2. Trạng thái hiện tại (Snapshot)

- **Cập nhật lần cuối**: `2026-11-21`
- **Phiên bản hiện có**: **v3.2.0 RELEASE COMPLETE (AutoCAD MCP Live Parser + 3-Zone Visual Beam Section Diagram + Git Backup & Rules Complete)**
  - **Kho mã nguồn Git**: `https://github.com/magicxlll/zOfit.git` (Nhánh `main` - tự động sao lưu & đồng bộ đa máy).
  - **1-Click Installer**: `dist/setup/zOfit_QS_Pro_Setup.exe` (`137.04 MB` / `143,694,535 bytes`, .NET 8 WinForms single-file compressed, nhúng sẵn `.xlam` + `zofit-engine.exe`, hỗ trợ cài mới / nâng cấp / gỡ cài đặt + `--silent`). *Lưu ý: File setup > 100MB được loại trừ khỏi Git theo `.gitignore` để tuân thủ hạn mức của GitHub; khi clone về máy mới có thể biên dịch lại bằng `build/build_release.ps1`.*
  - **Excel Add-in**: `dist/zOfit_QS_Pro.xlam` (18 VBA modules Pure-ASCII gồm `modCadParser.bas` & `modRebarConfig.bas` v3.2 + `customUI14.xml` Ribbon v3.2 có nút **`Quét Dầm CAD (MCP)`** kết nối cổng 9877). File add-in được theo dõi trực tiếp trên Git repo.
  - **Local Engine (.NET 8 Windows)**: `dist/engine/zofit-engine.exe` (`net8.0-windows` commands: `ping`, `open_rebar_editor`, `generate_structure`, `expand_rebar_presets`, `parse_cad`, `generate_finishing`, `generate_infra`, `parse_beam_section`, `scan_cad_beam`).
- **Tình trạng kiểm thử mới nhất (`2026-11-21`)**:
  - `dotnet test`: **22/22 PASS** (0 warnings, 0 errors — unit test cho `Tcvn3Converter`, `CadBeamSectionParser` gồm cả dầm 1 lớp thép và dầm lớn nhiều lớp `1300x1000` với 4 nhánh đai đứng + 2 nhánh đai ngang từ bản vẽ thực tế `S05-WATER TREATMENT.dwg`).
  - AutoCAD MCP Live Integration: Quét trực tiếp qua TCP Port 9877 vào AutoCAD đang mở bản vẽ S05 → Nhận diện chính xác 6 dầm (`1MBX1`, `1MBX9`, `1MBX2`, `1MBX4`, `1MBX6`, `1MBX8`) và dầm lớn (`1MBX3 / 1MBX5 / 1MBX7`).
  - Hình dạng thép & đai: T1/B1 nhận diện D07 (U móc 35D), tăng cường gối D06 (L neo 35D), đai ngoài D39, đai trong C-tie D24 (4 nhánh đứng + 2 nhánh ngang), thép giá D01, thép bên D01.
  - `build/test_xlam.ps1` + `build/test_v2_e2e.ps1`: **PASS 100%** (0 lỗi công thức trên 16 sheets xuất sạch).
- **Quy tắc phiên làm việc**:
  - Khi bắt đầu phiên trên bất kỳ máy nào: Tự động chạy `git pull origin main` (hoặc `clone` nếu máy mới) trước khi đọc `memory.md` và code.
  - Sau mỗi phiên debug/sửa lỗi: Tự động chạy unit test, cập nhật `memory.md`, commit và push lên `https://github.com/magicxlll/zOfit.git`.
- **Lưu ý môi trường**:
  - Máy người dùng có Excel với `DecimalSeparator = ,` / `ThousandsSeparator = .` → đã xử lý chuẩn hóa bằng `CultureInfo.InvariantCulture` (C#) và `Replace(..., ",", ".")` (VBA).

---

## 3. Bản đồ Kho mã nguồn & Tài liệu Tham chiếu

### Tài liệu điều hành dự án
- `memory.md` — Bộ nhớ liên phiên (file này)
- `docs/spec.md` — Đặc tả kỹ thuật & nghiệp vụ đầy đủ cho Wave 1 (v2.0) & Wave 2 (v2.1)
- `tasks/plan.md` — Kiến trúc, đồ thị phụ thuộc, kế hoạch 2 đợt & quản trị rủi ro
- `tasks/todo.md` — Bảng tracking tiến độ 21/21 task

### Mã nguồn v2.2 Complete
- `engine/zOfit.Engine/`: `Program.cs`, `Models/IpcModels.cs`, `Generators/RcDetailGenerator.cs`, `Generators/BarShapeCatalog.cs`, `Generators/RebarBbsGenerator.cs`, `Generators/FinishingEngine.cs`, `Generators/InfraEngine.cs`, `Sync/SyncEngine.cs`, `Parsers/CadTextParser.cs`
- `engine/zOfit.Engine.Tests/`: `PingTests.cs`, `RcDetailGeneratorTests.cs`, `RebarBbsGeneratorTests.cs`, `SyncEngineTests.cs`, `CadTextParserTests.cs`, `FinishingAndInfraTests.cs`
- `installer/zOfit.Installer/`: `zOfit.Installer.csproj`, `Program.cs`, `InstallService.cs`, `MainForm.cs`
- `src/vba/` (Pure-ASCII): `modUnicode.bas`, `modRibbon.bas`, `modSettings.bas`, `modCatalog.bas`, `modTools.bas`, `modShapeDraw.bas`, `modRebarTKT.bas`, `modSteelShape.bas`, `modRebarSummary.bas`, `modEngineBridge.bas`, `modSpecSheet.bas`, `modQsData.bas`, `modRebarConfig.bas`, `modQsGenerator.bas`, `modCadParser.bas`, `modBoqLinker.bas`, `modFinishingData.bas`, `modInfraData.bas`, `clsShapeBtn.cls`, `frmSteelLookup_code.txt`, `frmBarShape_code.txt`, `frmBoqLinker_code.txt`, `ThisWorkbook_code.txt`
- `src/ribbon/customUI14.xml` — Ribbon UI v2.1
- `src/data/`: `catalog_rebar.csv`, `catalog_materials.csv`, `catalog_shapes.csv`, `i18n.csv`, `shapes_bmp/D01.bmp..D41.bmp`
- `build/`: `build_release.ps1`, `build_xlam.ps1`, `test_xlam.ps1`, `test_v2_e2e.ps1`, `test_golden_bhflex.ps1`, `inject_ribbon.py`, `extract_catalog.py`, `prepare_shape_bmps.ps1`
- `dist/`: `setup/zOfit_QS_Pro_Setup.exe`, `zOfit_QS_Pro.xlam`, `engine/zofit-engine.exe`

---

## 4. Các Quyết định Kiến trúc quan trọng (ADR Log)

| ID | Quyết định | Lý do |
|---|---|---|
| **ADR-01** | **Kiến trúc Lai (`.xlam` + `zofit-engine.exe` .NET 8)** | Tách logic phức tạp sang C# dễ test xUnit, chạy cực nhanh, giữ 100% trải nghiệm trong Excel. |
| **ADR-02** | **Giao tiếp JSON File IPC qua `%TEMP%\zofit_ipc\`** | Tránh lỗi COM registration / `regasm` / xung đột Excel 32-64 bit / quyền Admin. |
| **ADR-03** | **Đầu ra thuần công thức Excel 2007–365** | Gửi file `.xlsx` cho Sếp/CĐT audit 100% minh bạch không cần Add-in. |
| **ADR-04** | **1-Click Installer `.NET 8 WinForms` single-file compressed `.exe`** | Nhúng sẵn `.xlam` + `zofit-engine.exe` (`EnableCompressionInSingleFile` giảm xuống `102.16 MB`), không cần Python/.NET Runtime/Inno Setup. |
| **ADR-05** | **Tương thích `Linkcad` & Non-Destructive Sync** | Ô kích thước trên `QS_Data` luôn mở; `Sync` bám `ElementID` và bảo toàn 100% dòng `MANUAL`. |

---

## 5. Bài học kỹ thuật & Bẫy đã xử lý (KHÔNG lặp lại!)

1. **VBA Pure ASCII**: Mọi `.bas`/`.cls` phải thuần ASCII (script `build_xlam.ps1` tự kiểm tra byte > 127); tiêu đề UserForm dùng `SetFormCaptionU`.
2. **VBA giới hạn độ dài dòng (~1023 ký tự)**: Dòng code vượt ~1023 chars (như nối chuỗi JSON dài 1 dòng) làm VBE hỏng ngầm khi `SaveAs` → báo lỗi mù mờ `Document not saved.`. Luôn chia `s = s & ...` thành nhiều dòng ngắn!
3. **Build `.xlam` khi đã cài add-in vào `%APPDATA%`**: Phải đóng mọi workbook `*zOfit*` trước khi `Workbooks.Open` / `SaveAs` và build qua tên file tạm `_build_<guid>.xlam` rồi `Move-Item`.
4. **VBA JSON Parser & `Scripting.Dictionary` Case Sensitivity**: `CreateObject("Scripting.Dictionary")` mặc định `vbBinaryCompare`! Phải set **`dict.CompareMode = vbTextCompare`** và dùng `Sub ParseValueInto(ByRef outVal)` + `Set outVal = Nothing` để tránh lỗi `Object required` (424) / treo modal error dialog.
5. **Không bao giờ gọi `MsgBoxU` trong `ErrHandler` khi `showMsg = False`**: Excel chạy COM ẩn sẽ treo vĩnh viễn nếu bật MsgBox ẩn!
6. **Ribbon XML**: Bên trong `<menu>` **BẮT BUỘC dùng `<menuSeparator />`**, escape `&lt;`/`&gt;`.
7. **Đăng ký Excel Add-in Registry (`OPEN`) & PowerShell `WinExe`**: Registry `HKCU\...\Excel\Options` `OPEN*` dùng chuẩn `"zOfit_QS_Pro.xlam"` + Trusted Location `%APPDATA%\Microsoft\AddIns\`. Khi test `.exe` WinExe từ PowerShell phải dùng `Start-Process -Wait -PassThru` mới đợi đồng bộ và đọc đúng `ExitCode`.
8. **Locale dấu phẩy thập phân (`DecimalSeparator = ,`)**: Mọi JSON truyền qua engine dùng `InvariantCulture` (dấu chấm `.`) và công thức COM dùng chuẩn US-English.

---

## 6. Nhật ký Phiên làm việc (Session Log)

| Phiên | Ngày | Tóm tắt công việc đã hoàn thành | Kết quả / Verify |
|---|---|---|---|
| **S01** | 2026-11-19 | Hoàn thành v1.0 Phase 1 & Phase 2 (`QS_Settings`, 1.559 vật liệu, `TKT Pro` 5.000 dòng, 41 hình dạng BMP, Light Mode, Live Preview, Coupler, Cắt tối ưu 1D). | `zOfit_QS_Pro.xlam` v1.0 PASS |
| **S02** | 2026-11-19 | Phân tích 7 file Spec & file thực tế `BOM_BHFLEX...xlsx`. Tạo `docs/spec.md`, `tasks/plan.md`, `tasks/todo.md`, `memory.md`. Khởi chạy `/goal` triển khai Engine `.NET 8` (15/15 xUnit PASS), 8 VBA modules v2.1, `frmBoqLinker`, `zOfit.Installer`, `build_xlam.ps1` v2.1. | Engine 15/15 PASS; Installer compiled |
| **S03** | 2026-11-19..20 | Debug triệt để 4 lỗi COM/VBA/Installer: (1) dòng VBA >1023 ký tự, (2) JSON Parser `Set` + `CompareMode = vbTextCompare`, (3) `ErrHandler` modal hang khi `showMsg=False`, (4) nén single-file `.exe` (`219MB -> 102.16MB`) & Registry `OPEN` `"zOfit_QS_Pro.xlam"`. Chạy PASS toàn bộ suite kiểm thử + Golden BHFLEX + UIAutomation trên Excel thật. | **RELEASE v2.1 COMPLETE (21/21 Tasks PASS 100%)** |
| **S04** | 2026-11-20 | Nâng cấp lớn theo phản hồi test thực tế của User (**v2.2.0**): (1) Thiết lập BT Lót riêng theo từng loại cấu kiện (`SPEC_Lean_Ftg/Beam/Slab/Tank_Thk & _Ext`) + ghi đè từng cấu kiện trên `QS_Data!P:Q`; (2) Kiến trúc Master-Detail `QS_Data` (19 cột `A:S`) ↔ `QS_RebarConfig` (16 cột `A:P`) hỗ trợ Dầm (trên/dưới/giữa/gia cường gối-nhịp), Sàn (1-2 phương, 1-2 lớp, mũ+con kê), Vách (thanh đứng rải L, ngang rải H, đai C), Cột (nối sole `SPEC_Col_StaggerOffset`), đai thông minh + chèn thanh `CUSTOM` tự do + nút `frmBarShape` 41 hình thép trực quan ngay trên `QS_RebarConfig`; (3) Sửa ưu tiên `FindEngineExecutable()` dùng bản cạnh `.xlam` trước `%LOCALAPPDATA%`. | `dotnet test` **18/18 PASS**, `test_xlam` + `test_v2_e2e` (16 sheets, 0 lỗi) + `test_golden_bhflex` + `Setup.exe` (`102.18 MB`) PASS 100% |
| **S05** | 2026-11-20 | Thiết kế lại toàn diện bảng thép **`QS_RebarConfig` v3.0 (Transparent Parametric BBS)** học theo tinh hoa file `AAA_BOQ_CAL RC WORK_2026_V2.xlsx`: (1) Bố cục 3 vùng 39 cột `A..AM`: Vùng A link sống từ `QS_Data`, Vùng B tham số tường minh (`BarDir`, `DistDir`, `HookL <50 => xD`, `Bv`, `HookR`, `@1/@2`, `SL.m`, `L.m`, `±Pb`, `±Lth`), Vùng C công thức Excel 2007–365 live ngay trên dòng (`Left..Mid..Right`, `QT'Y bar`, `Lbar`, `Lap`, `Sum Len`, `Density`, `Rebar kg`), chia theo Block cấu kiện + `SUBTOTAL(9)` tổng từng Block & `AL2`; (2) Ánh xạ 1-1 toàn bộ 41 dáng thép `D01..D41` sang `TKT_*`; (3) Sửa bẫy VBA `Dim grpCol As New Collection` trong vòng lặp + mở rộng menu chèn nhanh (`Chèn Thanh Chính / Đai / CUSTOM`). | `dotnet test` **18/18 PASS**, `test_v2_e2e` (`PC1_Blk=231.94kg`, `GBX1=7`, `C1=5 Stagger=600mm`, `Custom D17 Lbar=1800mm`) + `test_golden_bhflex` + `Setup.exe` (`102.19 MB`) PASS 100% |
| **S06** | 2026-11-20 | Xây dựng **Bảng Vẽ & Nhập Thép Trực Quan (`RebarEditorForm` + `StructuralDiagramRenderer` + `RebarParamCalculator`)** trong `zOfit.Engine` (v3.1) + Tài liệu [HUONG_DAN_QS_REBARCONFIG.md](docs/HUONG_DAN_QS_REBARCONFIG.md): (1) Chuột phải bất kỳ cấu kiện nào trên `QS_Data` hoặc `QS_RebarConfig` (hoặc nút Ribbon `Nhập Thép (Hình vẽ)`) → mở Form đồ họa vẽ mặt bằng/mặt cắt cấu kiện kèm ghi chú thanh thép, tô sáng thanh đang chọn, nút thêm nhanh tiếng Việt theo vị trí, bảng chỉnh sửa không giới hạn & nạp lại chỉnh tiếp bất kỳ lúc nào; (2) Giải quyết bẫy `STARTUPINFO.wShowWindow = SW_HIDE` bằng `ShowWindow(Handle, 5 /*SW_SHOW*/)` trên `RebarEditorForm.Load` + giữ `<OutputType>Exe</OutputType>` đồng bộ CLI. | `dotnet test` **19/19 PASS**, UIAutomation `-WindowStyle Hidden` xác nhận Form `1240x770` + `test_xlam` + `test_v2_e2e` + `test_golden_bhflex` + `Setup.exe` (`137.02 MB`) PASS 100% |
| **S07** | 2026-11-21 | Triển khai tính năng **Quét Dầm AutoCAD MCP (3 bước)**: (1) Nâng cấp `StructuralDiagramRenderer` vẽ mặt cắt 3 vùng Gối Trái - Giữa Nhịp - Gối Phải kèm trắc diện dầm, chỉ số hàm lượng thép $\mu$, thể tích BT, diện tích ván khuôn, khối lượng thép sống; (2) Xây dựng `AutoCadMcpClient` (kết nối trực tiếp TCP 9877), bộ chuyển đổi TCVN3 sang Unicode `Tcvn3Converter`, parser bảng dầm `CadBeamSectionParser`, lệnh CLI `parse_beam_section` & `scan_cad_beam`; (3) Tích hợp Ribbon button `btnScanCadBeam`, `modRibbon.bas`, `modCadParser.ScanCadBeamSectionToSheet`, `modRebarConfig.AppendOrUpdateRebarBlock` ghi dữ liệu thuần công thức Excel vào `QS_Data` & `QS_RebarConfig`. Test live với AutoCAD trích xuất 6 dầm thành công. | `dotnet test` **21/21 PASS**, AutoCAD MCP Live scan 6 dầm (`1MBX1`, `1MBX9`,...), `test_xlam` + `test_v2_e2e` PASS 100% (0 lỗi), `Setup.exe` (`137.04 MB`) |
| **S08** | 2026-11-21 | Hoàn thiện nhận diện dầm đa lớp & đai phức hợp + Sao lưu Git & Thiết lập Rules đa máy: (1) Nâng cấp `CadBeamSectionParser.cs` chuẩn hóa hình dạng thép: T1/B1 dáng D07 (U móc 35D), tăng cường gối dáng D06 (L neo 35D), đai ngoài D39, đai lồng trong C-tie D24 (4 nhánh đứng + 2 nhánh ngang B,C,D), thép giá D01, thép bên D01; (2) Thu nhỏ ngưỡng phân nhóm dòng từ 250 xuống 50 đơn vị để tách biệt rõ hàng Thép Trên và Thép Dưới; (3) Loại bỏ fallback tự quét khi chưa chọn đối tượng CAD; (4) Nâng cấp đồ họa mặt cắt dầm trực quan (`StructuralDiagramRenderer.cs`) thể hiện 2 lớp thép gối và hệ đai đa nhánh; (5) Làm rõ cơ chế nhận diện CAD: 100% C# .NET 8 local offline, 0% AI / GPU / Cloud; (6) Khởi tạo Git repo, cấu hình `.gitignore` loại trừ file nặng >100MB, backup toàn bộ lên `https://github.com/magicxlll/zOfit.git`, tạo rules (`GEMINI.md`, `AGENTS.md`, `.agents/rules/git_sync_rules.md`) tự động pull khi mở phiên và push sau mỗi lần debug. | `dotnet test` **22/22 PASS**, `build_release.ps1` PASS 100%, Git push `main` thành công lên GitHub |
