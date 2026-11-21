# zOfit QS Pro — Task Breakdown & Progress Tracker

> **Quy ước trạng thái**: `[ ]` Chưa làm · `[/]` Đang làm · `[x]` Hoàn thành & đã Verify · `[!]` Bị chặn
> **Quy tắc cập nhật**: Sau khi hoàn thành mỗi task (đã chạy Verify PASS), đánh dấu `[x]`, cập nhật **Bảng Tiến Độ Tổng Quan** và ghi chú vào [memory.md](../memory.md).

---

## 📊 Bảng Tiến Độ Tổng Quan (Progress Dashboard)

| Đợt (Wave) | Giai đoạn (Phase) | Số Task | Đã xong | Tiến độ | Trạng thái |
|---|---|---:|---:|---:|---|
| **v1.0 (Baseline)** | Phase 1 & Phase 2 (Settings, Catalog, TKT Pro 41 Shapes, Tools) | 12 | 12 | 100% | ✅ Hoàn tất & Verify |
| **Wave 1 (v2.0)** | Phase 1.1 — Foundation, IPC & Single-Entry Schemas (`W1-T01..04`) | 4 | 4 | 100% | ✅ Hoàn tất & Verify |
| **Wave 1 (v2.0)** | Phase 1.2 — Detail/TKT Generators & Non-Destructive Sync (`W1-T05..08`) | 4 | 4 | 100% | ✅ Hoàn tất & Verify |
| **Wave 1 (v2.0)** | Phase 1.3 — CAD Parser, BOQ Linker & 1-Click Installer (`W1-T09..13`) | 5 | 5 | 100% | ✅ Hoàn tất & Verify |
| **Wave 2 (v2.1)** | Phase 2.1 — Finishing Module (`W2-T01..03`) | 3 | 3 | 100% | ✅ Hoàn tất & Verify |
| **Wave 2 (v2.1)** | Phase 2.2 — Infrastructure Module (`W2-T04..06`) | 3 | 3 | 100% | ✅ Hoàn tất & Verify |
| **Wave 2 (v2.1)** | Phase 2.3 — Ribbon v2.1, Rebuild Installer & Full Regression (`W2-T07..08`) | 2 | 2 | 100% | ✅ Hoàn tất & Verify |
| **Wave 3 (v2.2)** | Phase 3 — Per-Type/Element Lean Concrete + Master-Detail `QS_RebarConfig` & Visual Shape Picker (`W3-T01..04`) | 4 | 4 | 100% | ✅ Hoàn tất & Verify |
| **Wave 4 (v3.0)** | Phase 4 — Transparent Parametric BBS `QS_RebarConfig` v3.0 (3-Zone 39-Col, Block Subtotals, 1-to-1 TKT Map `W4-T01..04`) | 4 | 4 | 100% | ✅ Hoàn tất & Verify |
| **Wave 5 (v3.1)** | Phase 5 — Visual Structural Diagram Rebar Editor (`RebarEditorForm` WinForms + GDI+ Blueprint, Right-Click `QS_Data`, Docs `W5-T01..03`) | 3 | 3 | 100% | ✅ Hoàn tất & Verify |
| **Wave 6 (v3.2)** | Phase 6 — AutoCAD MCP Live Integration, 3-Zone Beam Section Diagram & Git Cloud Sync Rules (`W6-T01..04`) | 4 | 4 | 100% | ✅ Hoàn tất & Verify |
| **TỔNG CỘNG (v2.0..v3.2)** | **6 Waves · 10 Phases · 10 Checkpoints** | **36** | **36** | **100%** | 🎉 **RELEASE v3.2 READY** |

---

# WAVE 1 — v2.0 (Core Hybrid Engine, Structure/Earthwork/TKT, CAD Parser, BOQ Linker & Installer)

## Phase 1.1: Foundation, IPC & Single-Entry Schemas

### - [x] `W1-T01`: Khởi tạo Solution `.NET 8` Engine + Hợp đồng JSON IPC + xUnit Test Harness
- **Verify**: `dotnet test` + `dotnet publish` (`dist/engine/zofit-engine.exe`) PASS.

### - [x] `W1-T02`: Xây dựng cầu nối VBA ↔ Engine (`modEngineBridge.bas`)
- **Verify**: `TestEnginePing` từ Excel VBA qua JSON UTF-8 (`"Đài cọc PC1 - Bê tông M400 Ø25"`) PASS; Pure-VBA recursive parser dùng `CompareMode = vbTextCompare` & `ParseValueInto` ByRef an toàn 32/64-bit.

### - [x] `W1-T03`: Xây dựng Sheet `QS_Spec` (Tiêu chuẩn Dự án từ 6 file Spec + Named Ranges `SPEC_*`)
- **Verify**: `modSpecSheet.bas` tạo đủ 39 Named Ranges `SPEC_*`, tương thích `CFG_*` cũ, serialize đầy đủ sang JSON.

### - [x] `W1-T04`: Xây dựng Sheet `QS_Data` (Bảng Nhập Cấu Kiện 1 Lần — Single-Entry & Tương thích `Linkcad`)
- **Verify**: `modQsData.bas` mở khóa ô hình học `F..O` cho `Linkcad`, tự sinh `ElementID`, validate trùng lặp/thiếu trường.

### 🚦 Checkpoint `CP-1.1` — ✅ PASSED

---

## Phase 1.2: Core Formula Generators & Non-Destructive Sync

### - [x] `W1-T05`: Engine sinh `<Building>_Detail` (Đất, Cọc PHC, Bê tông lót, Bê tông kết cấu, Cốp pha)
- **Verify**: `RcDetailGeneratorTests` & Excel COM test PASS, 0 `#REF!`, toàn bộ công thức trỏ thẳng về `QS_Data!` & `SPEC_*`.

### - [x] `W1-T06`: Engine sinh `TKT_<Building>` (Thống Kê Thép Chuẩn Tự Động Link về `QS_Data` & `QS_Spec`)
- **Verify**: `RebarBbsGeneratorTests` kiểm tra đủ 8 `RebarTemplate` (`FTG_2WAY`, `BEAM_STD`, `COL_STD`, `SLAB_2WAY_2LAYER`, `WALL_2FACE`, `STAIR_6BAR`, `MANHOLE_STD`, `PILE_HEAD`) + tự đổi Coupler theo `SPEC_CouplerMinD` + tóm tắt theo `BOQ_Bucket`.

### - [x] `W1-T07`: Engine Đồng bộ không phá hủy (`SyncEngine.cs`)
- **Verify**: `SyncEngineTests` & Excel COM test PASS: bảo toàn 100% dòng `MANUAL`, thêm `PC7`, cập nhật `PC1`, đánh dấu `[REMOVED]` cho `PC2`.

### - [x] `W1-T08`: VBA Bulk Sheet Writer, Outline Grouping & Summary Bucket (`modQsGenerator.bas`)
- **Verify**: Sinh `49.ELECTRIC SUPPLY_Detail` + `TKT_49.ELECTRIC SUPPLY` + `TH_Thep_ToanCT` (< 2 giây), đổi `QS_Data!G5` nhảy số đồng loạt qua `CalculateFullRebuild`.

### 🚦 Checkpoint `CP-1.2` — ✅ PASSED

---

## Phase 1.3: CAD Raw Parser, BOQ Linker & 1-Click Installer

### - [x] `W1-T09`: Bộ tách dữ liệu CAD thô (`CadTextParser.cs` + `modCadParser.bas`)
- **Verify**: Tách chính xác **58/58** dòng text CAD thực tế trích từ sheet `Data` của `BOM_BHFLEX...xlsx` + Block Attribute `TN-TENGA2` -> `QS_CAD_Parsed` (`SUMIFS`/`AVERAGEIF`).

### - [x] `W1-T10`: Công cụ Map-Link an toàn vào BOQ Chủ đầu tư (`frmBoqLinker` + `modBoqLinker.bas`)
- **Verify**: Thu thập đủ nguồn `*_Detail` + `TH_Thep_ToanCT` theo `BOQ_Bucket`, gán công thức link vào BOQ đích không phá format.

### - [x] `W1-T11`: Cập nhật Ribbon v2.1 (`customUI14.xml`), `modRibbon.bas` & Pure-ASCII Check
- **Verify**: Valid OOXML, UIAutomation xác nhận tab `zOfit QS Pro` hiển thị trực tiếp trên Microsoft Excel thật.

### - [x] `W1-T12`: Đóng gói Bộ cài đặt 1-Click (`zOfit_QS_Pro_Setup.exe` — `.NET 8 WinForms`)
- **Verify**: `dist/setup/zOfit_QS_Pro_Setup.exe` (`102.16 MB` single-file compressed), chạy `--silent` giải nén khớp SHA256 100% `.xlam` & `zofit-engine.exe` + ghi Registry `OPEN` & Trusted Location.

### - [x] `W1-T13`: Golden Benchmark Test trên dữ liệu thực tế BHFLEX Structure
- **Verify**: `build/test_golden_bhflex.ps1` PASS 100%, Export `.xlsx` sạch 0 macro, 0 lỗi công thức.

### 🚦 Checkpoint `CP-1.3` — ✅ RELEASE v2.0 VERIFIED

---

# WAVE 2 — v2.1 (Finishing & Infrastructure Modules + Full E2E Regression)

## Phase 2.1: Finishing Module (Hoàn thiện)

### - [x] `W2-T01`: Template `QS_DoorWin` & `QS_Rooms` (`modFinishingData.bas`)
### - [x] `W2-T02`: Engine Hoàn thiện (`FinishingEngine.cs`) — Xây/Trát/Sơn tự trừ cửa qua `INDEX/MATCH`, Chống thấm WC (+300mm), Dây chuyền Mái R1
### - [x] `W2-T03`: Tích hợp VBA Hoàn thiện & kiểm chứng thay đổi kích thước cửa `1SD1` tự nhảy khối lượng `Masonry/Plaster`

### 🚦 Checkpoint `CP-2.1` — ✅ PASSED

## Phase 2.2: Infrastructure Module (Hạ tầng)

### - [x] `W2-T04`: Template `QS_Infra` (`modInfraData.bas`) — Đường, Cống/Mương, Tường rào
### - [x] `W2-T05`: Engine Hạ tầng (`InfraEngine.cs`) — Đào taluy, Đệm, Trừ chiếm chỗ ống `PI()*(D/2)^2*L`, BTXM + Dowel, Bó vỉa/Terrazzo, Tường rào
### - [x] `W2-T06`: Tích hợp VBA Hạ tầng & Nạp chiều dài tự động từ `QS_CAD_Parsed` (`D200 = 44m`)

### 🚦 Checkpoint `CP-2.2` — ✅ PASSED

## Phase 2.3: Ribbon v2.1, Rebuild Installer & Full BHFLEX Regression

### - [x] `W2-T07`: Build & Package `zOfit_QS_Pro_Setup.exe` v2.1 (`build/build_release.ps1`)
### - [x] `W2-T08`: Full End-to-End Regression (`dotnet test` 15/15 + `test_xlam.ps1` + `test_v2_e2e.ps1` 0 lỗi công thức + `test_golden_bhflex.ps1` + UIAutomation Ribbon PASS)

### 🚦 Checkpoint `CP-2.3` — ✅ RELEASE v2.1 COMPLETE (21/21 Tasks)

---

# WAVE 3 — v2.2 (Tùy biến BT Lót đa cấu kiện + Kiến trúc Thép Master-Detail `QS_RebarConfig`)

### - [x] `W3-T01`: Bê tông lót theo từng loại cấu kiện (`SPEC_Lean_Ftg/Beam/Slab/Tank_*`) & ghi đè từng cấu kiện (`QS_Data!P:Q`)
### - [x] `W3-T02`: Kiến trúc 2 bảng `QS_Data` (`A:S`) ↔ `QS_RebarConfig` (`A:P`) + Bung Preset tự động (`expand_rebar_presets`)
### - [x] `W3-T03`: Tích hợp `frmBarShape` (41 hình thép `D01..D41`) & chèn thanh `CUSTOM` trực tiếp trên `QS_RebarConfig` + Ribbon mới
### - [x] `W3-T04`: Rebuild `zOfit_QS_Pro_Setup.exe` v2.2 & Full Regression (`dotnet test` 18/18, `test_xlam`, `test_v2_e2e` 16 sheets, `test_golden_bhflex` + UIAutomation PASS)

### 🚦 Checkpoint `CP-3.1` — ✅ RELEASE v2.2 COMPLETE (25/25 Tasks)

---

# WAVE 4 — v3.0 (Transparent Parametric BBS `QS_RebarConfig` phong cách `AAA_BOQ_CAL` + Link 1-1 sang `TKT_*`)

### - [x] `W4-T01`: Thiết kế schema 3 vùng 39 cột `A..AM` trên `QS_RebarConfig` (`modRebarConfig.bas`), công thức live 100% Excel 2007–365 (không `LET`/`IFS`), Block Header + `SUBTOTAL(9)` từng cấu kiện
### - [x] `W4-T02`: Nâng cấp `IpcModels.cs` & `RebarBbsGenerator.cs` sinh 11 Preset kèm đầy đủ `BarDir`/`DistDir`/`HookL/Bv/HookR`/`@1/@2`/`±Lth` & ánh xạ 1-1 toàn bộ 41 hình `D01..D41` sang `TKT_*`
### - [x] `W4-T03`: Bổ sung nút Ribbon & Right-Click (`Chèn Thanh Chính`, `Chèn Thanh Đai`, `Chèn CUSTOM`, `Chọn 41 Dáng Thép` + highlight ô `Le.1a..Ri.2a`)
### - [x] `W4-T04`: Full Regression & Đóng gói `.exe` (`dotnet test` 18/18, `test_xlam`, `test_v2_e2e` 16 sheets 0 lỗi, `test_golden_bhflex` + UIAutomation PASS)

### 🚦 Checkpoint `CP-4.1` — ✅ RELEASE v3.0 COMPLETE (29/29 Tasks)

---

# WAVE 5 — v3.1 (Bảng Vẽ Nhập Thép Trực Quan theo Cấu Kiện & Hướng dẫn chi tiết)

### - [x] `W5-T01`: Xây dựng `RebarParamCalculator.cs`, `StructuralDiagramRenderer.cs` (GDI+ mặt bằng/mặt cắt 6 nhóm cấu kiện + tô sáng thanh chọn) & `RebarEditorForm.cs`
### - [x] `W5-T02`: Tích hợp menu Chuột phải `QS_Data` & `QS_RebarConfig`, nút Ribbon **`Nhập Thép (Hình vẽ)`**, ghi/cập nhật Block không giới hạn (`ReplaceOrAppendElementBlock`)
### - [x] `W5-T03`: Kiểm thử UIAutomation (`-WindowStyle Hidden`), `dotnet test` 19/19 PASS, `build_release.ps1` + `test_golden_bhflex.ps1` + tài liệu [HUONG_DAN_QS_REBARCONFIG.md](../docs/HUONG_DAN_QS_REBARCONFIG.md)

### 🚦 Checkpoint `CP-5.1` — ✅ RELEASE v3.1 COMPLETE (32/32 Tasks)

---

# WAVE 6 — v3.2 (AutoCAD MCP Live Integration, 3-Zone Beam Section Diagram, Multi-Layer Rebar & Git Cloud Sync)

### - [x] `W6-T01`: Xây dựng cầu nối AutoCAD MCP trực tiếp qua TCP Port 9877 (`AutoCadMcpClient.cs`), bộ chuyển đổi font TCVN3 sang Unicode (`Tcvn3Converter.cs`)
### - [x] `W6-T02`: Xây dựng Parser bảng dầm kết cấu AutoCAD (`CadBeamSectionParser.cs`) hỗ trợ dầm đa lớp thép (T1/B1 D07, gia cường gối D06, đai ngoài D39, đai lồng C-tie D24 đa nhánh ngang/dọc, thép bên D01, thép giá D01) + Unit tests 22/22 PASS
### - [x] `W6-T03`: Nâng cấp biểu đồ mặt cắt dầm trực quan (`StructuralDiagramRenderer.cs` + `RebarEditorForm.cs`) 3 vùng (Gối Trái - Giữa Nhịp - Gối Phải), hiển thị đầy đủ các lớp chấm thép và các nhánh đai; tích hợp Ribbon button `btnScanCadBeam` và macro `ScanCadBeamSectionToSheet`
### - [x] `W6-T04`: Khởi tạo Git repository, cấu hình `.gitignore` loại trừ file nặng >100MB, backup toàn bộ lên `https://github.com/magicxlll/zOfit.git`, thiết lập `GEMINI.md`, `AGENTS.md` và `.agents/rules/git_sync_rules.md` tự động pull khi mở phiên và push sau mỗi lần debug

### 🚦 Checkpoint `CP-6.1` — ✅ RELEASE v3.2 COMPLETE (36/36 Tasks)

