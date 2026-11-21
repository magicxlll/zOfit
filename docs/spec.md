# Spec: zOfit QS Pro — Hệ thống Bóc Tách Khối Lượng Bán Tự Động (v2.0 & v2.1)

> **Trạng thái**: Draft v1.0 — Chờ xác nhận lần cuối trước khi triển khai Wave 1 (`W1-T01`)
> **Ngày tạo**: 2026-11-19
> **Tham chiếu thực tế**: `BOM_BHFLEX_CONSTRUCTION_Sub building 49-50 & Civil_VIETPH_261119.xlsx`, `AAA_BOQ_CAL RC WORK_2026_V2.xlsx`, `TKT.xlsm`, `Spec & form/*`

---

## 1. Objective (Mục tiêu & Ý định cốt lõi)

### 1.1 Bài toán thực tế cần giải quyết
Trong quy trình bóc tách dự thầu hàng ngày (nhà xưởng công nghiệp FDI / hạ tầng như BHFLEX V2F):
1. **Nhập trùng lặp**: Cùng một cấu kiện (móng `PC1`, dầm `GBX1`, sàn `GS1`...) phải nhập tay 2–3 lần: nhập ở `*_Detail` để tính Đất/Bê tông/Ván khuôn, rồi nhập lại ở sheet `RC` để tính Cốt thép.
2. **Link thủ công dễ vỡ**: Copy/sửa cụm công thức thủ công giữa Bê tông → Bê tông lót → Ván khuôn → Đào đất phát sinh lỗi `#REF!` (đã thấy 9 lỗi `#REF!` ngay trong file thực tế `49.ELECTRIC SUPPLY_Detail` và `50.WWT_Detail`).
3. **Chuẩn hóa Spec thủ công**: Các quy chuẩn kỹ thuật dự án (lớp bảo vệ, neo/nối thép, đổi Coupler, trừ uốn, taluy hố đào, hao hụt, chiều dày lót/mở rộng) đang bị rải rác hardcode trong công thức từng ô.
4. **Dữ liệu CAD thô**: Bảng thống kê Text/Block Attribute xuất từ AutoCAD vào sheet `Data` (`A-W300- i0.2%-L=13m`, `Box:600x600-L11m`, `D200-L44m`, `TN-TENGA2`) phải xử lý tách chuỗi thủ công.
5. **File nặng do rác**: File copy qua nhiều dự án bị phình tới `10.000–42.000` Defined Names rác.

### 1.2 Mục tiêu sản phẩm (North Star)
Xây dựng **zOfit QS Pro** — hệ thống QS bán tự động theo **Kiến trúc Lai (Excel Add-in + Local Compiled Engine)**, đóng gói thành **1 file cài đặt `.exe` 1-click**:
- **100% làm việc ngay trong Microsoft Excel** (tab Ribbon `zOfit QS Pro`), không tạo phần mềm/môi trường làm việc tách biệt.
- **Nhập 1 lần (Single-Entry)** kích thước + thông số cấu kiện trên `QS_Data` (phân theo `Khu vực / Nhà`) + tiêu chuẩn dự án trên `QS_Spec` → **1-click sinh ra toàn bộ bản tính chi tiết thuần công thức Excel sống**:
  - `*_Detail` (format CĐT: Đào/Lấp đất, Cọc, Bê tông lót, Bê tông kết cấu, Cốp pha)
  - `TKT_*` (Bảng thống kê thép chuẩn có 41 hình dạng thép, link trực tiếp kích thước từ `QS_Data` & quy chuẩn từ `QS_Spec`)
  - Link khối lượng sang BOQ Tổng (`2. ARC+STRUC.`, `3. CIVIL& LANDSCAPE`) giữ nguyên cấu trúc BOQ Chủ đầu tư.
- **Sản phẩm đầu ra là file `.xlsx` thuần công thức Excel 2007–365**: Sếp, Tư vấn, Chủ đầu tư mở xem và kiểm tra từng công thức mà **không cần cài Add-in**.

### 1.3 Ranh giới Tự động hóa vs. Bán thủ công (Không "ảo hóa")

| Nhóm hạng mục | Tự động 90–100% (Engine sinh công thức sống) | Bán tự động (Template chuẩn + Công thức) | Thủ công `DxRxC` (Giữ nguyên tự do chỉnh sửa) |
|---|---|---|---|
| **Kết cấu & Đất** *(Wave 1)* | Từ `QS_Data` + `QS_Spec` → sinh `*_Detail` (Đào máy/thủ công, Lấp đất/cát, Vận chuyển, Đầm nền, Cấp phối, Màng PE, BT lót, BT kết cấu, Cốp pha trừ giao chuẩn) + `TKT_*` (số thanh theo `@`, neo/nối, bảo vệ, Coupler, hình dạng) + Link BOQ. | Cọc PHC (khoan dẫn, ép âm, mối nối, thép đầu cọc, nén tĩnh); cột `Trừ giao bổ sung / Hệ số` cho bể ngầm, dầm vát. | Cấu kiện dị hình bất quy tắc (móng máy phức tạp, ram xoắn): chèn dòng `DxRxC` trực tiếp trong cụm `SUBTOTAL(9,...)` và được giữ nguyên khi `Sync`. |
| **Dữ liệu CAD thô** *(Wave 1)* | Bộ tách `CAD Text/Block Parser` đọc vùng text/block dump từ `Data` → trích `Loại / Đường kính / Kích thước / Độ dốc / L / Cao độ`. | Ánh xạ nhanh sang `QS_Data` hoặc `QS_Infra`; tương thích hoàn toàn ô mở với công cụ `Linkcad` hiện có của người dùng. | Đo `dim` / vẽ `pline` trên AutoCAD (nằm ngoài phạm vi v2.x, dự kiến giai đoạn sau). |
| **Hoàn thiện** *(Wave 2)* | Bảng `QS_DoorWin` (cửa đi/cửa sổ) → mọi dòng tường xây/trát/sơn dùng `INDEX/MATCH` tự trừ cửa; dây chuyền Mái (`S` mái link 5–6 lớp). | Bảng `QS_Rooms` (`Phòng | P | S | H | Mã cửa`) → sinh Xây, Trát, Sơn, Lát, Ốp, Len, Trần, Chống thấm WC (+vén chân 300mm). | Chi tiết kiến trúc lẻ (lan can, phào chỉ, biển báo, thiết bị vệ sinh, vách đặc thù): nhập `DxRxC` ngay trên BOQ tổng. |
| **Hạ tầng** *(Wave 2)* | Tự lấy `L` từ `CAD Parser` cho tuyến cống/mương/ống; Hố ga & Bể ngầm dùng chung `QS_Data` (BT/VK/Thép). | Tuyến ống/mương (Đào taluy, Đệm, Trừ chiếm chỗ ống, Đắp trả, Vận chuyển); Đường nội bộ (K95/K98, Đá dăm, BTN, Bó vỉa, Vỉa hè). | San nền lưới ô vuông/Civil3D, cảnh quan cây xanh, đấu nối đặc thù. |
| **BOQ Tổng** *(Wave 1+2)* | Dọn rác Defined Names, Kiểm lỗi `#REF!/#VALUE!`, Map-Link công thức vào ô khối lượng BOQ CĐT (giữ nguyên `GAP A-B`). | Tạo khung BOQ chuẩn khi chưa có file CĐT. | Không ghi đè định dạng/cấu trúc file BOQ mẫu của CĐT. |

---

## 2. Architecture & Tech Stack (Kiến trúc Lai)

### 2.1 Sơ đồ Kiến trúc Lai
```
┌──────────────────────────────────────────────────────────────────────┐
│ Microsoft Excel (64-bit / 32-bit, Excel 2016–365)                    │
│  └─ zOfit_QS_Pro.xlam (Thin Ribbon UI + Sheet Bridge + Live Picker)  │
│       • customUI14.xml (Tab "zOfit QS Pro")                          │
│       • Pure-ASCII VBA + Win32 Unicode + 41 Embedded Shape BMPs      │
│       • Đọc range ActiveWorkbook ──► %TEMP%\zofit_ipc\req_<id>.json  │
│       • Gọi zofit-engine.exe (ẩn console, đồng bộ)                   │
│       • Nhận res_<id>.json ──► ghi bulk array (1-shot) vào Excel     │
└───────────────────────────────┬──────────────────────────────────────┘
                                │ JSON IPC (File-based, Zero COM/Reg)
┌───────────────────────────────▼──────────────────────────────────────┐
│ %LOCALAPPDATA%\zOfitQSPro\zofit-engine.exe                           │
│  • .NET 8 (C#) Self-Contained Single-File (win-x64), KHÔNG cần cài   │
│    Python/.NET Runtime trên máy người dùng                           │
│  • Modules:                                                          │
│     - SpecEngine: Chuẩn hóa rule từ 6 file spec                      │
│     - RcDetailGenerator: Sinh ma trận dòng + công thức *_Detail      │
│     - RebarBbsGenerator: Sinh công thức TKT_* link về QS_Data        │
│     - SyncEngine: Đồng bộ không phá hủy theo ElementID               │
│     - CadTextParser: Regex tách text/block CAD (ống, mương, hố ga)   │
│     - BoqLinker: Gợi ý & sinh công thức link sang BOQ CĐT            │
│     - FinishingEngine & InfraEngine (Wave 2)                         │
└──────────────────────────────────────────────────────────────────────┘
┌──────────────────────────────────────────────────────────────────────┐
│ Phân phối: dist\zOfit_QS_Pro_Setup.exe (.NET 8 WinForms Single-File) │
│  • Nhúng sẵn zOfit_QS_Pro.xlam + zofit-engine.exe bên trong          │
│  • 1-click Cài đặt / Nâng cấp / Gỡ bỏ:                               │
│     1) Kiểm tra & nhắc đóng Excel nếu đang mở                        │
│     2) Copy .xlam → %APPDATA%\Microsoft\AddIns\ (gỡ Zone.Identifier) │
│     3) Copy zofit-engine.exe → %LOCALAPPDATA%\zOfitQSPro\            │
│     4) Đăng ký HKCU\...\Excel\Options (OPEN / OPENx) tự bật Add-in   │
└──────────────────────────────────────────────────────────────────────┘
```

### 2.2 Tại sao chọn Kiến trúc này?
- **Không dùng COM-visible `.dll` / `.xll` phức tạp**: Tránh lỗi đăng ký `regasm`, xung đột Excel 32/64-bit, và quyền Admin IT doanh nghiệp.
- **JSON File IPC + Bulk Variant Array write trong VBA**: C# xử lý 5.000 cấu kiện trong `< 150ms`; VBA ghi mảng 2D `ws.Range(...).Formula = arr` trong `< 300ms` — tổng thời gian sinh toàn bộ `*_Detail` + `TKT_*` dưới 1 giây.
- **Thuần công thức Excel chuẩn (Excel 2007–365)**: Tuyệt đối không dùng `_xlfn.LET`, `_xlfn.IFS`, `XLOOKUP` hay UDF riêng của Add-in → người nhận mở trên mọi phiên bản Excel đều thấy đúng số và audit được từng công thức.

### 2.3 Tech Stack đã xác minh trên máy build
- **Excel Add-in**: VBA + OpenXML `customUI14.xml` (build qua `build/build_xlam.ps1` + `build/inject_ribbon.py`, Python 3.11.9).
- **Local Engine**: C# / **.NET 8.0 SDK** (`8.0.428` đã có sẵn), publish `-r win-x64 --self-contained true -p:PublishSingleFile=true`, test bằng **xUnit**.
- **1-Click Installer**: C# / **.NET 8.0 WinForms** single-file `.exe` (không cần Inno Setup hay NSIS).

---

## 3. Data Model & Sheet Specifications

### 3.1 Sheet `QS_Spec` (Tiêu chuẩn Kỹ thuật Dự án — Chuẩn hóa từ 6 file Spec)
Mỗi thông số đều có **Named Range phạm vi Workbook (`SPEC_*`)** để mọi công thức ở `*_Detail`, `TKT_*`, `QS_Rooms`, `QS_Infra` tham chiếu tường minh. Hỗ trợ 3 Preset (`BHFLEX / FDI Industrial`, `TCVN / TT13-2021`, `Custom`).

| Nhóm Spec | Thông số & Named Range tiêu biểu | Giá trị mặc định (theo 6 file Spec) |
|---|---|---|
| **A. Cao độ & Đào đắp** | `SPEC_Elev_Natural` (Cao độ tự nhiên), `SPEC_Elev_FFL` (Cao độ ±0.000), `SPEC_Topsoil_Thk` (Bóc hữu cơ `0.10m`), `SPEC_Exc_WorkSpace_Ftg` (`0.70m`), `SPEC_Exc_WorkSpace_Beam` (`0.50m`), `SPEC_Exc_SlopeK` (Hệ số mái dốc/taluy `1.30`), `SPEC_Exc_MachinePct` (`0.90`), `SPEC_Exc_LaborPct` (`0.10`), `SPEC_Backfill_CompK` (`1.30`) | Theo `49_Detail` & `Spec ha tang.rtf` |
| **B. Bê tông & Ván khuôn** | `SPEC_Lean_Thk` (`0.10m`), `SPEC_Lean_Overhang` (`0.10m`), `SPEC_Lean_Slab_Thk` (`0.05m`), `SPEC_Conc_Waste` (`1.03`), `SPEC_PE_Overlap` (`1.01`), `SPEC_Opening_DeductMin` (`0.50 m²` hoặc `1.00 m²` theo spec) | Theo `BIM BOQ Engine.md` & `Logic Đo Bóc Xây Dựng.md` |
| **C. Quy chuẩn Cốt thép** | Lớp bảo vệ (`SPEC_Cover_Ftg=50mm`, `SPEC_Cover_Col=35mm`, `SPEC_Cover_Beam=30mm`, `SPEC_Cover_Slab=20mm`, `SPEC_Cover_Wall=30mm`, công thức `MAX(c_base, d)`), Hệ số neo (`SPEC_Lan_Tension=40d`, `SPEC_Lan_Comp=30d`), Hệ số nối (`SPEC_Llap_Mult=45d`), Ngưỡng Coupler (`SPEC_CouplerMinD=20`), Chiều dài cây chuẩn (`SPEC_StdBarLen=11.7m`), Trừ/cộng uốn (`45°=1d, 90°=2d, 135°=3d, 180°=+6d`), Hao hụt thép (`SPEC_Scrap_Small=3% [<=10]`, `SPEC_Scrap_Med=2.5% [12-18]`, `SPEC_Scrap_Large=2% [>18]`) | Theo `Quy chuẩn thép.md` |
| **D. Cọc PHC & Nền xưởng** | `SPEC_Pile_CutHead` (`0.50m`), `SPEC_Pile_PlugLen` (`1.50m`), `SPEC_Joint_Grid` (`6.0m`), `SPEC_Slab_Waste` (`1.03`) | Theo `Logic Đo Bóc Xây Dựng.md` |
| **E. Hoàn thiện & Hạ tầng** *(Wave 2)* | `SPEC_Wall_OpeningMin` (`0.50m²`), `SPEC_Mesh_Width` (`0.20m`), `SPEC_WP_CurbH` (`0.30m`), `SPEC_Pipe_Bedding` (`0.15m`), `SPEC_Pipe_Slope` (`0.50`) | Theo `spec hoan thien.rtf` & `Spec ha tang.rtf` |

### 3.2 Sheet `QS_Data` (Bảng Dữ Liệu Cấu Kiện Nhập 1 Lần — Single-Entry)
- **Tương thích `Linkcad`**: Các cột kích thước (`G..N`) là ô nhập số thuần (không khóa, không đè công thức nếu user chấm CAD trực tiếp vào ô).
- **Cấu trúc cột chính (A–AA)**:
  - **Định danh & Phân loại (`A–E`)**: `ElementID` (khóa duy nhất để Sync, vd `49-FTG-PC1`), `Building/Zone` (vd `49.ELECTRIC SUPPLY`, `50.WWT`, `CIVIL`), `Group` (`Foundation`, `GroundBeam`, `Column`, `Wall`, `Beam`, `Slab`, `Stair`, `Lintel`, `Pile`, `Manhole`, `Tank`), `Mark` (`PC1`, `GBX1`, `C1`, `GS1`...), `Description` (Tên cấu kiện).
  - **Hình học (`F–N`)**: `Qty` (Số lượng CK), `L (m)`, `W (m)`, `H/t (m)`, `TopElev (m)`, `BotElev (m)`, `ConcGrade` (`M400`, `M350`...), `ClashDeduct_V (m³)` (trừ giao BT nếu có), `ClashDeduct_FW (m²)` (trừ giao VK nếu có), `IsBuried` (`Y/N` — tự đưa vào Đào đất & BT lót).
  - **Thông số Thép nhanh (`O–Z`)** — để sinh tự động các dòng `TKT_*` link ngược về `L, W, H`:
    - `RebarTemplate` (mẫu bố trí theo loại cấu kiện: `FTG_2WAY`, `BEAM_STD`, `COL_STD`, `SLAB_2WAY_2LAYER`, `WALL_2FACE`, `STAIR_6BAR`, `MANHOLE_STD`, `CUSTOM`)
    - `Main_X_Dia`, `Main_X_Spec` (số thanh hoặc `@200`), `Main_Y_Dia`, `Main_Y_Spec`
    - `Top_X_Dia`, `Top_X_Spec`, `Top_Y_Dia`, `Top_Y_Spec`
    - `Stirrup_Dia`, `Stirrup_Spec` (vd `100/200`), `Extra_Spec` (thép gia cường / đai C / giằng)
  - **Trạng thái (`AA`)**: `BOQ_Bucket` (`Substructure`, `Superstructure`, `Misc/Stair/Lintel`, `Civil`).

### 3.3 Sheet `<Building>_Detail` (Format Bóc Tách CĐT — Wave 1)
Tự sinh 1 sheet cho mỗi `Building/Zone` trong `QS_Data` (ví dụ `49.ELECTRIC SUPPLY_Detail`, `50.WWT_Detail`), khớp cấu trúc BHFLEX:
- **Cột**: `A: Code/ID | B: Description | C: Spec/Dim (=F*1000&"x"&...) | D: Unit | E: NO | F: L | G: W | H: H/t | I: a | J: b | K: c | L: QTY (=PRODUCT(E:K) hoặc công thức chu vi cốp pha) | M: ElementID (ẩn/nhạt để Sync)`.
- **Cấu trúc phân cấp (`SUBTOTAL(9,...)` + Excel Outline Grouping 3 cấp)**:
  1. **Pile Works** (Khoan dẫn, Ép cọc, Ép âm, Mối nối, Cắt đầu cọc, Bịt đầu cọc, Vận chuyển)
  2. **Soil Works** (Đào máy 90%, Đào thủ công 10%, Lấp đất tận dụng, Lấp cát mua mới, Vận chuyển đất thừa, Đầm đất K95, Cấp phối đá dăm K98, Màng PE 2 lớp) — tất cả `L, W, H, NO` link trực tiếp bằng công thức từ `QS_Data` + `QS_Spec`!
  3. **Formwork Works** (VK bê tông lót, VK Móng/Dầm móng/Sàn trệt, VK Cột/Vách/Dầm/Sàn thân, VK Bổ trụ/Lanh tô/Cầu thang)
  4. **Concrete Works** (BT lót M100, BT Móng/Sàn trệt/Ram, BT Thân, BT Bổ trụ/Lanh tô/Cầu thang, Đánh bóng mặt sàn, Cắt khe co giãn)
- **Đảm bảo Zero `#REF!`**: Công thức trong `*_Detail` trỏ thẳng về `QS_Data!<Row>` và `SPEC_*` (không link vòng vèo ngược lên-xuống trong cùng sheet như cách làm tay cũ!).

### 3.4 Sheet `TKT_<Building>` (Thống Kê Thép Chuẩn Tự Động Link — Wave 1)
Kế thừa toàn bộ thành quả Phase 2 (`modRebarTKT.bas`, 41 hình dạng `D01..D41`, Light Mode 0-Shape, Live Preview, Coupler tự động, 5.000 dòng) và nâng cấp:
- Sinh tự động các dòng thanh thép theo `RebarTemplate` của từng `ElementID` trong `QS_Data`:
  - **Ví dụ Móng `PC1` (`FTG_2WAY`)**: Sinh thanh thép lớp dưới phương X (`D02`, `L1 = QS_Data!G*1000 - 2*MAX(SPEC_Cover_Ftg, Ø)`, `L2 = QS_Data!I*1000 - 2*MAX(...)`, `Thanh/CK = ROUNDUP((QS_Data!H*1000 - 2*c)/@ + 1, 0)`, `Số CK = QS_Data!F`) và phương Y tương tự.
  - **Ví dụ Sàn `GS1` (`SLAB_2WAY_2LAYER`)**: Sinh 4 dòng (Dưới X, Dưới Y, Trên X, Trên Y) + thép mũ/gia cường; đổi kích thước `A, B, h` ở `QS_Data` → chiều dài & số lượng thanh rải nhảy tức thì!
  - **Ví dụ Cột/Dầm/Vách/Cầu thang (`ST-01..ST-06`)**: Sinh đủ thép chủ (tự cộng neo `SPEC_Lan`, nối `SPEC_Llap` hoặc chuyển Coupler khi `Ø >= SPEC_CouplerMinD`) và thép đai (tự tính chu vi theo `W, H` trừ bảo vệ, số đai theo bước `@`).
- Tự động tổng hợp sang `TH_Thep_ToanCT` chia đúng 3 nhóm BOQ (`Substructure`, `Superstructure`, `Misc/Stair/Lintel`) theo `BOQ_Bucket` + tự tính hao hụt theo đường kính (`<=10: 3%`, `12–18: 2.5%`, `>18: 2%`).

### 3.5 Non-Destructive ID-Based Sync Protocol
Khi người dùng thêm/bớt cấu kiện trong `QS_Data` và bấm **Sync (Đồng bộ)**:
1. Engine đọc cột `ElementID` hiện có trong `*_Detail` và `TKT_*`.
2. **Dòng do người dùng tự chèn tay** (không có `ElementID` hoặc đánh dấu `MANUAL` ở cột tag) → **GIỮ NGUYÊN 100%** tại đúng cụm `SUBTOTAL(9,...)`.
3. **Dòng `ElementID` cũ còn tồn tại** → Cập nhật công thức nếu `RebarTemplate`/`Group` thay đổi (không xóa định dạng/ghi chú tay).
4. **Dòng `ElementID` mới** → Chèn đúng vào cụm tương ứng và tự co giãn phạm vi `SUBTOTAL(9,...)`.
5. **Dòng `ElementID` đã bị xóa khỏi `QS_Data`** → Đánh dấu highlight vàng `[REMOVED]` để QS tự kiểm tra xóa (không âm thầm xóa dữ liệu gây mất đối chiếu).

### 3.6 CAD Raw Text & Block Parser (`Data` → Bảng sạch — Wave 1)
Đọc vùng dữ liệu raw CAD dump (giống sheet `Data` của BHFLEX):
- Nhận diện tự động các mẫu text thực tế:
  - Mương/Cống hộp: `A-W300- i0.2%-L=13m`, `B-W400- L=6m`, `Box:600x600-L11m-i0.2%`
  - Ống tròn: `D200-L44m- i0.5%`, `HDPE D300 L=25.5m`
  - Block Attribute: `TEN BLOCK | MG | DC | SO LUONG` (tính `H_đào = MG - DC`, nhóm theo loại hố ga)
- Xuất ra bảng `QS_CAD_Parsed` (thuần Excel `SUMIFS` / `AVERAGEIFS`) sẵn sàng bấm 1 nút đẩy sang `QS_Data` hoặc `QS_Infra`.

### 3.7 BOQ Auto-Linker (Wave 1)
- Giao diện `frmBoqLinker` trong Excel:
  - Bên trái: Danh sách các đầu mục tổng hợp có sẵn từ `*_Detail` và `TH_Thep_ToanCT` (hiển thị tên hạng mục, đơn vị, giá trị hiện tại, địa chỉ ô nguồn).
  - Bên phải: Chọn ô đích trên BOQ Chủ đầu tư (`2. ARC+STRUC.` hoặc `3. CIVIL& LANDSCAPE`) + hệ số quy đổi đơn vị tùy chọn (`kg -> tấn = /1000`, hao hụt `*1.03`).
  - Bấm **Gán Link**: Ghi công thức trực tiếp vào cột `G` (hoặc ô đang chọn) mà **không làm hỏng định dạng/cột `GAP (A-B, %)` của BOQ CĐT**.

### 3.8 Wave 2 Modules (v2.1 — Hoàn thiện & Hạ tầng)
1. **Hoàn thiện (`QS_DoorWin` + `QS_Rooms` → `Finishing_Detail`)**:
   - `QS_DoorWin`: Danh mục cửa đi/cửa sổ/vách kính (`Code | W | H | S=W*H | Lintel_L=W+0.5`).
   - `QS_Rooms`: Bảng phòng & tường (`Building | Floor | Room | Area | Perim | Height | WallType | WallLen | DeductDoorCodes | FloorFinish | WallFinish | CeilingType | IsWetArea`).
   - Tự sinh công thức sống: Xây tường (trừ cửa qua `INDEX/MATCH`), Lanh tô + lưới chống nứt 200mm, Trát, Sơn, Lát nền, Ốp tường, Len chân tường, Chống thấm WC (`S + P*SPEC_WP_CurbH`), Trần thạch cao, và **Dây chuyền Mái (`R1`)** (1 ô `S_mái` link 5–6 lớp chống thấm/tạo dốc/lưới thép/đánh bóng).
2. **Hạ tầng (`QS_Infra` → `Civil_Detail`)**:
   - **Tuyến cống/mương**: Link `L` từ `QS_CAD_Parsed` + cao độ đáy/mặt đất → tự sinh Đào đất taluy, Đệm cát/đá, Lắp đặt ống, Đắp trả trừ thể tích chiếm chỗ ống, Vải địa kỹ thuật, Vận chuyển đất thừa.
   - **Đường nội bộ & Vỉa hè**: Nhập diện tích/chiều dài từng loại đường (`Road_Type | Area | Length | Width`) → tự sinh Đào khuôn, K95, K98, Đá dăm, BTN C12.5/BTXM + khe co giãn + thép dowel, Bó vỉa, Vỉa hè Terrazzo.
   - **Hố ga, Tường rào, Bể ngầm/WWTP**: Hố ga & Bể ngầm tận dụng chung `QS_Data` (ra đủ Đào đất chiếm chỗ + BT + VK + Thép).

---

## 4. Commands (Lệnh Build, Test & Package)

```powershell
# 1. Build & Test Local Engine (.NET 8)
dotnet test engine/zOfit.Engine.Tests/zOfit.Engine.Tests.csproj --verbosity normal
dotnet publish engine/zOfit.Engine/zOfit.Engine.csproj -c Release -r win-x64 --self-contained true -p:PublishSingleFile=true -o dist/engine

# 2. Build Excel Add-in (.xlam) + Inject Ribbon + Embedded 41 BMPs
powershell -ExecutionPolicy Bypass -File build/build_xlam.ps1

# 3. Run Excel COM & UIAutomation Integration Tests
powershell -ExecutionPolicy Bypass -File build/test_xlam.ps1

# 4. Build 1-Click Installer (.exe) embedding .xlam + zofit-engine.exe
dotnet publish installer/zOfit.Installer/zOfit.Installer.csproj -c Release -r win-x64 --self-contained true -p:PublishSingleFile=true -o dist/setup
```

---

## 5. Project Structure (Cấu trúc thư mục)

```text
D:\03_APP\3. System\GITHUB\zOfit\
├── memory.md                       # Bộ nhớ cục bộ lâu dài cho AI Agent (cập nhật mỗi phiên)
├── docs/
│   └── spec.md                     # Tài liệu đặc tả kỹ thuật & nghiệp vụ (file này)
├── tasks/
│   ├── plan.md                     # Kế hoạch triển khai 2 đợt & phân tích rủi ro
│   └── todo.md                     # Danh sách Task chi tiết & bảng Tracking tiến độ
├── src/
│   ├── vba/                        # Pure-ASCII VBA modules (.bas, .cls, *_code.txt)
│   ├── ribbon/customUI14.xml       # Ribbon XML cho tab zOfit QS Pro
│   └── data/                       # CSV catalogs (rebar, materials, 41 shapes, i18n, presets) + shapes_bmp/
├── engine/                         # .NET 8 Local Compiled Engine (C#)
│   ├── zOfit.Engine/               # CLI + JSON IPC + QS Generators + CAD Parser + Sync
│   └── zOfit.Engine.Tests/         # xUnit unit & golden-file tests
├── installer/
│   └── zOfit.Installer/            # .NET 8 WinForms 1-click Setup.exe
├── build/                          # PowerShell & Python build/test scripts
├── dist/                           # Build outputs (zOfit_QS_Pro.xlam, engine/, setup/zOfit_QS_Pro_Setup.exe)
├── Spec & form/                    # 7 file Spec & BOQ mẫu gốc của người dùng (Read-only)
└── Hinhdangthep/                   # 41 ảnh gốc dang (1..41).png (Read-only)
```

---

## 6. Code Style & Technical Conventions

1. **VBA (`src/vba/`)**:
   - **Bắt buộc Pure ASCII**: Không gõ tiếng Việt có dấu trực tiếp trong `.bas`/`.cls`. Dùng `T("key")` (tra từ `DB_I18N`) hoặc `U("\uXXXX")`. Title bar UserForm dùng `SetFormCaptionU Me, U(...)`.
   - **Không dùng `Dim x As New Cls` trong vòng lặp**: Luôn khai báo `Dim x As Cls` rồi `Set x = New Cls`.
   - **Công thức Excel 2007–365**: Cấm `_xlfn.LET`, `_xlfn.IFS`, `XLOOKUP`, `FILTER`, `UNIQUE`. Dùng `IF`, `INDEX/MATCH`, `SUMIF/SUMIFS`, `PRODUCT`, `SUBTOTAL(9,...)`, `ROUNDUP`.
   - **Ribbon XML (`customUI14.xml`)**: Trong `<menu>` bắt buộc dùng `<menuSeparator />` (không dùng `<separator />`); escape `<` thành `&lt;`, `>` thành `&gt;`.
2. **C# Engine (`engine/`)**:
   - Pure functions cho mọi bộ sinh công thức (Input: DTO JSON → Output: SheetGrid DTO JSON) để unit test 100% không cần mở Excel.
   - Mọi hàm sinh công thức phải tự kiểm tra không phát sinh `#REF!`, không dùng ký tự locale dấu phẩy thập phân (luôn `CultureInfo.InvariantCulture`).

```csharp
// Ví dụ chuẩn sinh công thức thuần Excel trong RebarBbsGenerator.cs
// Thanh lớp dưới phương X cho móng PC1 tại dòng rData trên QS_Data:
string l1Formula = $"=QS_Data!G{rData}*1000-2*MAX(SPEC_Cover_Ftg,O{rTkt})";
string qtyPerEl  = $"=ROUNDUP((QS_Data!H{rData}*1000-2*MAX(SPEC_Cover_Ftg,O{rTkt}))/{spacingMm},0)+1";
```

---

## 7. Testing Strategy (Chiến lược Kiểm thử)

1. **Unit Tests (`dotnet test` — xUnit)**:
   - Kiểm tra công thức sinh ra cho từng loại cấu kiện (`Footing`, `GroundBeam`, `Column`, `Wall`, `Beam`, `Slab`, `Stair`, `Pile`, `Manhole`, `Room`, `Pipe`, `Road`).
   - Kiểm tra `CadTextParser` với 100% chuỗi thật lấy từ sheet `Data` của file BHFLEX.
   - Kiểm tra `SyncEngine` bảo toàn dòng `MANUAL` khi thêm/sửa/xóa `ElementID`.
2. **Excel COM Integration Tests (`build/test_xlam.ps1`)**:
   - Khởi động Excel thật qua COM, nạp `dist/zOfit_QS_Pro.xlam`, chạy quy trình end-to-end (tạo `QS_Spec` → điền `QS_Data` → gọi `zofit-engine.exe` → sinh `*_Detail` + `TKT_*` → `CalculateFullRebuild`).
   - Assert **0 ô lỗi** (`#REF!`, `#NAME?`, `#VALUE!`, `#N/A`, `#DIV/0!`), kiểm tra số liệu khớp giá trị kỳ vọng (±0.1%), kiểm tra xuất file sạch `.xlsx` mở không cần Add-in.
3. **UIAutomation Ribbon Smoke Test**:
   - Mở `excel.exe` thực tế, kiểm tra tab `zOfit QS Pro` hiển thị trên thanh Ribbon.
4. **Golden Benchmark BHFLEX**:
   - Nhập thử cụm cấu kiện thực tế của `49.ELECTRIC SUPPLY` (`PC1..PC6, DR1, GBX1..GM2, GS1`) vào `QS_Data` → đối chiếu tổng Đào đất, BT lót, BT kết cấu, Cốp pha, Thép với file gốc `BOM_BHFLEX`.

---

## 8. Boundaries (3 Tầng Giới hạn)

- **Always do (Luôn làm)**:
  - Chạy `dotnet test` + `build_xlam.ps1` + `test_xlam.ps1` xác minh Pass sau mỗi task code.
  - Giữ tương thích Excel 2007–365 và giữ ô nhập kích thước mở cho `Linkcad`.
  - Cập nhật `tasks/todo.md` và `memory.md` sau mỗi phiên làm việc.
- **Ask first (Hỏi trước khi làm)**:
  - Thay đổi cấu trúc cột `QS_Data` hoặc tên Named Range `SPEC_*` sau khi đã chốt.
  - Thêm thư viện NuGet bên thứ ba nặng hoặc yêu cầu quyền Admin trong Installer.
  - Xóa/đổi tên các nút Ribbon cũ đang dùng ở v1.0.
- **Never do (Tuyệt đối không làm)**:
  - Không sửa trực tiếp các file gốc trong `Spec & form/`, `Hinhdangthep/`, `BOM_BHFLEX...xlsx`, `TKT.xlsm`.
  - Không tạo ứng dụng làm việc tách rời Excel hoặc xuất file `.xlsx` phụ thuộc macro/UDF mới xem được số.
  - Không để công thức hardcode spec (phải dẫn về `SPEC_*`) và không commit file tạm/scratch vào repo.

---

## 9. Success Criteria (Tiêu chí nghiệm thu)

### Wave 1 (v2.0)
- [ ] Cài đặt thành công trên máy sạch bằng **1 cú click `zOfit_QS_Pro_Setup.exe`** (không cần cài Python hay .NET Runtime), mở Excel thấy ngay tab `zOfit QS Pro`.
- [ ] Nhập 1 lần nhóm móng + dầm móng + sàn trệt nhà `49.ELECTRIC SUPPLY` vào `QS_Data` → sinh `49.ELECTRIC SUPPLY_Detail` + `TKT_49.ELECTRIC SUPPLY` trong **< 3 giây**, **0 lỗi công thức**.
- [ ] Đổi 1 kích thước `L` của `PC1` trên `QS_Data` → Đào đất, BT lót, BT móng, Cốp pha, chiều dài thanh thép & số lượng thanh thép trên `TKT_*`, và BOQ Tổng **tự cập nhật tức thì bằng công thức Excel** không cần chạy lại tool.
- [ ] Bấm `Sync` sau khi thêm `PC7` và chèn 2 dòng thủ công `DxRxC` → `PC7` tự vào đúng cụm `SUBTOTAL(9,...)`, **2 dòng thủ công giữ nguyên 100%**.
- [ ] `CAD Text/Block Parser` tách đúng 100% danh sách 53 dòng text & 35 dòng block mẫu từ sheet `Data` của file BHFLEX.
- [ ] Export `.xlsx` sạch: 0 Defined Names rác, 0 macro, mở trên máy không cài Add-in vẫn tính đúng 100%.

### Wave 2 (v2.1)
- [ ] `QS_DoorWin` + `QS_Rooms` tự sinh khối lượng Xây/Trát/Sơn (tự trừ cửa qua `INDEX/MATCH`), Lát/Ốp/Len/Trần/Chống thấm WC + Dây chuyền Mái `R1`.
- [ ] `QS_Infra` tự sinh Đường nội bộ, Tuyến ống/mương (link `L` từ `QS_CAD_Parsed`), Hố ga & Bể ngầm với 0 lỗi công thức.
- [ ] Chạy thử nghiệm End-to-End trên bộ dữ liệu mẫu BHFLEX (Kết cấu + Hoàn thiện + Hạ tầng) đạt khớp đối chiếu kiểm toán.
