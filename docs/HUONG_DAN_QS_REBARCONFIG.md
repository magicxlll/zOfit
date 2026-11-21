# Hướng Dẫn Sử Dụng Bảng Thép Tường Minh `QS_RebarConfig` (v3.0)

> **zOfit QS Pro v3.0 — Transparent Parametric BBS**
> Tài liệu hướng dẫn chi tiết cho Kỹ sư QS bóc tách và kiểm soát cốt thép trực tiếp ngay trên Excel theo tư duy minh bạch từng thanh thép (kết hợp mô hình nhập 1 lần `QS_Data` + bảng kiểm soát tham số trực tiếp như `AAA_BOQ_CAL`), xuất ra **100% công thức Excel 2007–365 sống**.

---

## 1. Tư duy cốt lõi & Luồng làm việc tổng thể

Thay vì phải gõ lại kích thước cấu kiện nhiều lần hay copy-paste hàng chục dòng công thức như cách làm truyền thống, **zOfit QS Pro** tổ chức dữ liệu theo quan hệ **Cha – Con (1-nhiều)**:

```text
[QS_Spec] (Tiêu chuẩn Dự án: Bv, Neo, Nối, Coupler, Sole cột, Hao hụt)
     +
[QS_Data] (Bảng nhập 1 lần: Kích thước LxWxH, Số lượng, Cao độ, Lót BT, RebarPreset)
     │
     ▼  (Bấm "Nạp / Cập nhật Block Thép từ QS_Data")
[QS_RebarConfig v3.0] ──► BẢNG NHẬP & KIỂM SOÁT THÉP CHÍNH (Hiển thị LIVE: Left..Mid..Right, SL thanh, Lbar, Lap, kg)
     │
     ▼  (Bấm "3. Sinh Detail + TKT")
[TKT_<Building>] + [<Building>_Detail] + [TH_Thep_ToanCT] ──► Map-Link vào BOQ CĐT
```

**Điểm mạnh của `QS_RebarConfig v3.0`**:
1. **Không bao giờ phải gõ lại Dài × Rộng × Cao**: Mọi thông số cấu kiện trên `QS_Data` tự link sang bằng `INDEX/MATCH`. Sửa `L/W/H` hoặc `Qty` bên `QS_Data` → chiều dài & khối lượng thép nhảy theo tức thì.
2. **Kiểm soát trực tiếp từng thanh ngay lúc nhập**: Nhìn từ trái sang phải thấy ngay từng đoạn `Left (L1) .. Mid (L4) .. Right (L7)`, số thanh `QT'Y bar`, tổng `Lbar (mm)`, đoạn nối `Lap (mm)` và khối lượng `Rebar (kg)` theo thời gian thực.
3. **Khớp 100% sang bảng Thống kê thép (`TKT_<Building>`)**: Toàn bộ 41 hình dạng `D01..D41` trên `TKT_*` link công thức 1-1 trực tiếp về `QS_RebarConfig`.

---

## 2. Quy trình 5 bước thao tác hàng ngày (Có 2 cách nhập siêu nhanh)

1. **Bước 1 — Thiết lập Spec (`QS_Spec`)**: Vào Tab Ribbon **`zOfit QS Pro`** → bấm **`Tiêu chuẩn QS_Spec`** để rà lớp bảo vệ (`SPEC_Cover_Ftg/Col/Beam/Slab/Wall`), hệ số neo/nối (`SPEC_Lan_Tension = 40D`, `SPEC_Llap_Mult = 45D`), khoảng lệch nối sole cột (`SPEC_Col_StaggerOffset = 600mm`), ngưỡng `SPEC_CouplerMinD` (vd `>= D20` dùng Coupler), chiều dài cây thép thương mại (`11700mm`).
2. **Bước 2 — Nhập cấu kiện 1 lần (`QS_Data`)**: Bấm **`1. Nhập QS_Data`**, nhập danh sách cấu kiện (hoặc đổ số liệu từ Linkcad vào các cột `F..O`), chọn **`RebarPreset`** ở cột `R` (`FTG_2WAY`, `BEAM_FULL`, `COL_STAGGER`, `WALL_2FACE`, `SLAB_2WAY_2LAYER`...).
3. **Bước 3 — Nhập & Sửa Thép trực quan (KHÔNG cần nhớ mã!)**:
   - **Cách A — Dùng Bảng Vẽ Cấu Kiện Trực Quan (Khuyên dùng cho mọi người mới & sửa nhanh)**:
     1. Chọn dòng cấu kiện bất kỳ ngay trên **`QS_Data`** (hoặc trong Block của cấu kiện trên **`QS_RebarConfig`**).
     2. Bấm **Chuột phải → `zOfit QS: Nhập / Sửa Thép Trực Quan (Hình vẽ)...`** (hoặc bấm nút **`Nhập Thép (Hình vẽ)`** trên Ribbon).
     3. Cửa sổ đồ họa mở ra hiển thị **bản vẽ mặt bằng & mặt cắt cấu kiện** kèm mũi tên chú thích vị trí từng thanh thép.
     4. Bấm các nút tiếng Việt bên trái (**`+ Đáy X`**, **`+ Chủ Trên`**, **`+ Gia Cường Gối`**, **`+ Đai Kín (@1/@2)`**, **`+ Đứng 2 Mặt`**, **`+ Chân Chó`**, **`+ Thanh Tự Do`**...) hoặc bấm chọn từng dòng trong bảng bên dưới (thanh đang chọn tự tô đỏ cam trên bản vẽ kèm hình dáng `D01..D41` & khối lượng `kg` tính ngay tức thì).
     5. Bấm **`✓ Đồng ý & Ghi vào Bảng Data`** → toàn bộ thanh thép & công thức Excel sống được ghi tự động vào đúng Block của cấu kiện trên **`QS_RebarConfig`**. Có thể mở lại để sửa/thêm **không giới hạn** bất cứ lúc nào!
   - **Cách B — Sửa hàng loạt trực tiếp trên Excel (`QS_RebarConfig`)**:
     - Bấm **`2. Bảng Thép QS_RebarConfig`** (hoặc **`Công cụ Thép CK → Nạp / Cập nhật Block Thép từ QS_Data`**) để sửa nhanh trực tiếp trên ô Excel.
4. **Bước 4 — Sinh bảng Chi tiết & Thống kê Thép (`Detail + TKT`)**:
   - Bấm **`3. Sinh Detail + TKT`** để xuất `<Building>_Detail`, `TKT_<Building>` (chuẩn 41 hình dạng thép) và `TH_Thep_ToanCT`.
5. **Bước 5 — Link BOQ & Xuất file sạch**:
   - Dùng **`3. Map-Link BOQ`** để gán công thức vào BOQ CĐT và **`Xuất .xlsx Sạch`** khi gửi hồ sơ kiểm toán.

---

## 3. Giải phẫu 39 Cột trên `QS_RebarConfig v3.0` (`A..AM`)

Bảng được cố định tiêu đề ở **Dòng 4** & **Cột D** (`Freeze Panes`), chia làm **3 Vùng màu rõ ràng**:

### Vùng A (`A..H` — Nền xanh xám): Liên kết tự động từ `QS_Data` *(Không cần sửa)*
| Cột | Tiêu đề | Ý nghĩa & Công thức |
|---|---|---|
| `A` | `ElementID` | Mã cấu kiện duy nhất (Dropdown từ `QS_Data!$A$5:$A$2000`) |
| `B` | `Building` | Tên nhà / khu vực (`INDEX/MATCH` từ `QS_Data!B`) |
| `C` | `Group` | Nhóm cấu kiện: Foundation, GroundBeam, Column, Wall, Beam, Slab... |
| `D` | `Mark` | Ký hiệu cấu kiện (PC1, GBX1, C1, W1...) |
| `E` | `QT'Y Ele` | Số cấu kiện giống nhau (`QS_Data!F`) |
| `F` | `X / L (mm)` | Chiều dài cấu kiện quy ra **mm** (`QS_Data!G * 1000`) |
| `G` | `Y / W (mm)` | Chiều rộng / bề dày cấu kiện quy ra **mm** (`QS_Data!H * 1000`) |
| `H` | `H (mm)` | Chiều cao / chiều dày sàn quy ra **mm** (`QS_Data!I * 1000`) |

### Vùng B (`I..Y` — Nền vàng nhạt): Tham số nhập Thép *(Khu vực thao tác chính)*
| Cột | Tiêu đề | Cách nhập thông minh |
|---|---|---|
| `I` | `BarID` | Số hiệu thanh thép (vd: `GBX1-T1`, `C1-1A`, `M1`...) |
| `J` | `Role / Vi tri` | Vai trò thanh (Dropdown: `BOT.X`, `TOP`, `SIDE`, `TOP-Goi`, `BOT-Nhip`, `MAIN.A`, `STIR.R`, `TIE.C`, `CHAIR`, `CUSTOM`...) |
| `K` | `Shape` | Mã dáng thép `D01..D41`. Chọn ô này rồi bấm **Chọn Dáng Thép** (trên Ribbon hoặc chuột phải) để chọn trực quan |
| `L` | `D (mm)` | Đường kính thép (`6, 8, 10, 12, 14, 16, 18, 20, 22, 25, 28, 32, 36, 40`) — Ô tô vàng nổi bật |
| `M` | `Part` | Hệ số mặt / phần giống nhau (VD: Thép vách 2 mặt nhập `2`; gia cường gối 2 đầu dầm nhập `2`; mặc định `1`) |
| `N` | **`BarDir`** | **Phương thanh chạy** — quyết định công thức tính đoạn giữa `Mid` *(Xem mục 4.2)* |
| `O` | **`DistDir`** | **Phương rải thanh** — quyết định công thức tính số lượng thanh `QT'Y bar` *(Xem mục 4.3)* |
| `P` | `HookL (xD\|mm)` | Chiều dài mỏ trái `Left`: nhập `< 50` tự nhân `×D`; nhập `>= 50` lấy thẳng `mm`; `0`/trống = không bẻ mỏ |
| `Q` | `Bv (mm)` | Lớp bê tông bảo vệ (mặc định link `=SPEC_Cover_*`, có thể gõ đè số riêng cho từng thanh) |
| `R` | `HookR (xD\|mm)` | Chiều dài mỏ phải `Right`: quy ước giống `HookL` |
| `S` | `@1 / @X` | Bước rải chính (`mm`), hoặc bước đai vùng gối/nối (`100`, `150`), hoặc bước lưới theo phương `X` (`600`) |
| `T` | `@2 / @Y` | Bước đai vùng giữa nhịp (`200`), hoặc bước lưới theo phương `Y`/`H` (`600`). Để trống = dùng chung `@1` |
| `U` | **`SL.m (Ovr)`** | **Ghi đè Số lượng thanh**: Gõ số vào → lấy thẳng số đó (vd `4` cây chủ dầm). Để trống → tự tính theo `DistDir` & `@` |
| `V` | **`L.m (Ovr)`** | **Ghi đè Đoạn giữa `Mid`**: Gõ số hoặc công thức vào → lấy thẳng chiều dài đó (`mm`). Để trống → tự tính theo `BarDir` |
| `W` | `+-Pb (mm)` | Cộng (+) hoặc Trừ (-) phạm vi rải thép (`mm`) trước khi chia `@` |
| `X` | `+-Lth (mm)` | Cộng (+) hoặc Trừ (-) chiều dài đoạn giữa `Mid` (`mm`) — VD cột nối sole gõ `=SPEC_Llap_Mult*L{r}+SPEC_Col_StaggerOffset` |
| `Y` | `Ghi chu` | Diễn giải chi tiết (tự động đưa sang cột Ghi chú và tên cấu kiện bên `TKT_*`) |

### Vùng C (`Z..AL` — Nền trắng/xanh) & Cột đánh dấu `AM`: Kết quả Công thức LIVE
| Cột | Tiêu đề | Công thức tự động (Excel 2007–365 thuần) |
|---|---|---|
| `Z` | `Left (L1)` | `=IF(P="",0,IF(P<50,L*P,P))` |
| `AA` | `Le.1a (L2)` | Ô mở: nhập thêm đoạn uốn phụ trái khi chọn dáng nhiều đoạn (vd `D17`, `D18`, `D22`...) |
| `AB` | `Le.1b (L3)` | Ô mở: nhập đoạn uốn phụ trái thứ 2 (khi chọn dáng >= 6 đoạn) |
| `AC` | `Mid (L4)` | Tự tính từ `(L.m, BarDir, X, Y, H, Bv, +-Lth)` |
| `AD` | `Ri.2b (L5)` | Ô mở: nhập đoạn uốn phụ phải thứ 2 |
| `AE` | `Ri.2a (L6)` | Ô mở: nhập đoạn uốn phụ phải khi chọn dáng nhiều đoạn |
| `AF` | `Right (L7)` | `=IF(R="",0,IF(R<50,L*R,R))` |
| `AG` | `QT'Y bar` | Số lượng thanh trên **1 `Part` của 1 cấu kiện**, tự tính từ `(SL.m, DistDir, X, Y, H, @1, @2, +-Pb)` |
| `AH` | `Lbar (mm)` | Tổng chiều dài hình học 1 thanh: `=SUM(Z:AF)` |
| `AI` | `Lap (mm)` | Đoạn nối chồng tự động: `=0` nếu `D < 10` hoặc `D >= SPEC_CouplerMinD`; ngược lại = `MAX(0, ROUNDUP(Lbar/SPEC_StdBarLen, 0) - 1) * D * SPEC_Llap_Mult` |
| `AJ` | `Sum Len (m)` | Tổng chiều dài quy mét: `=ROUND(QT'Y_Ele * Part * QT'Y_bar * (Lbar + Lap) / 1000, 2)` |
| `AK` | `Density` | Khối lượng riêng (kg/m): `=ROUND(PI() * (D/1000)^2 / 4 * 7850, 3)` |
| `AL` | **`Rebar (kg)`** | Khối lượng thép dòng: `=ROUND(SumLen * Density, 2)`. Dòng Block Header màu vàng dùng `=SUBTOTAL(9, ...)` cộng riêng từng cấu kiện; ô **`AL2`** cộng toàn bộ bảng không bị đếm trùng |
| `AM` | `RowKind` | Thẻ hệ thống (`BLK_HDR` / `BAR` / `BLK_END`) — **không nên tự sửa cột này** |

---

## 4. Bảng tra Quy tắc Tính Thông Minh

### 4.1. Quy ước Mỏ / Móc (`HookL` ở cột `P`, `HookR` ở cột `R`)
- Nhập **số `< 50`**: Hiểu là **hệ số `× D`**.
  - Ví dụ với `D16`: nhập `20` → `20 × 16 = 320 mm`; nhập `10` (móc đai `D10`) → `10 × 10 = 100 mm`; nhập `40` (neo `40D`) → `40 × 14 = 560 mm`.
- Nhập **số `>= 50`**: Hiểu là **kích thước thực tế (`mm`)**.
  - Ví dụ: nhập `150` → `150 mm`; nhập `300` → `300 mm`.
- Nhập **`0` hoặc để trống**: Không có mỏ (`0 mm`).

### 4.2. Danh sách `BarDir` (Cột `N` — Quyết định đoạn chính `Mid`)
| Mã `BarDir` | Công thức tính `Mid (mm)` (nếu `L.m` đang trống) | Thường dùng cho |
|---|---|---|
| `X` | `X - 2*Bv + (±Lth)` | Thép chạy dọc theo chiều dài `X / L` (móng, dầm, sàn, vách ngang) |
| `Y` | `Y - 2*Bv + (±Lth)` | Thép chạy theo chiều rộng `Y / W` (móng, sàn phương Y) |
| `H` | `H - 2*Bv + (±Lth)` | Thép đứng theo chiều cao `H` (cột, vách đứng, đai quây hông móng) |
| `X*0.25` | `ROUND(X*0.25, 0) + (±Lth)` | Thép mũ gối sàn / gia cường gối dầm (`L/4` mỗi bên, dùng kèm `Part=2`) |
| `Y*0.25` | `ROUND(Y*0.25, 0) + (±Lth)` | Thép mũ gối sàn theo phương `Y` |
| `X*0.7` | `ROUND(X*0.7, 0) + (±Lth)` | Thép gia cường bụng nhịp dầm (`0.7L`) |
| `PERIM_YH` | `2 * (Y + H - 4*Bv) + (±Lth)` | Chu vi đai kín dầm / giằng móng (tiết diện `W × H`) |
| `PERIM_XY` | `2 * (X + Y - 4*Bv) + (±Lth)` | Chu vi đai bao ngoài cột (tiết diện `X × Y`) |
| `PERIM_XY_INNER` | `2 * (ROUND(X*0.67,0) + Y - 4*Bv) + (±Lth)` | Đai lồng trong cột |
| `TIE_Y` | `Y - 2*Bv + (±Lth)` | Đai móc C kẹp theo phương rộng/bề dày `Y/W` (dầm, cột, vách) |
| `TIE_X` | `X - 2*Bv + (±Lth)` | Đai móc C kẹp theo phương `X/L` |
| `SLOPED_XH` | `ROUND(SQRT(X^2 + H^2), 0) + (±Lth)` | Thép bản thang nghiêng |
| `CHAIR_H` | `MAX(120, (H - 2*Bv)*2 + 150) + (±Lth)` | Chân chó kê thép sàn/móng (2 chân đứng + mặt trên) |
| `CUSTOM` | `0 + (±Lth)` *(kết hợp điền `L.m`)* | Thanh đặc biệt tự do hoàn toàn |

### 4.3. Danh sách `DistDir` (Cột `O` — Quyết định số thanh `QT'Y bar`)
| Mã `DistDir` | Công thức tính `QT'Y bar` (nếu `SL.m` đang trống) | Thường dùng cho |
|---|---|---|
| `FIXED` | `1` *(nhập số thanh vào cột `SL.m`)* | Thép chủ dầm, thép chủ cột, thanh gia cường cố định |
| `X` | `ROUND((X + (±Pb)) / @1, 0) + 1` | Thép rải dọc theo `X` (vd: thép móng phương Y, thép đứng vách) |
| `Y` | `ROUND((Y + (±Pb)) / @1, 0) + 1` | Thép rải dọc theo `Y` (vd: thép móng phương X, sàn phương X) |
| `H` | `ROUND((H + (±Pb)) / @1, 0) + 1` | Thép rải dọc theo chiều cao `H` (vd: thép ngang vách, đai quây móng) |
| `ZONE2_X` | `ROUND((X*0.5)/@1,0) + ROUND((X*0.5)/@2,0) + 2` | **Đai dầm 2 vùng**: nửa gối rải `@1`, nửa giữa rải `@2` |
| `ZONE2_H` | `ROUND((H*0.5)/@1,0) + ROUND((H*0.5)/@2,0) + 1` | **Đai cột 2 vùng**: vùng nối rải `@1`, vùng giữa rải `@2` |
| `XY_GRID` | `(ROUND(X/@1,0)+1) * (ROUND(Y/@2,0)+1)` | **Lưới `X × Y`** cho Chân chó kê thép sàn / móng (`@1 × @2`) |
| `XH_GRID` | `(ROUND(X/@1,0)+1) * (ROUND(H/@2,0)+1)` | **Lưới `X × H`** cho Đai C giữ 2 lớp thép vách (`@1 × @2`) |

---

## 5. Ví dụ thực chiến cho từng loại Cấu kiện

| Cấu kiện | Thao tác nhanh trên `QS_RebarConfig` |
|---|---|
| **Móng (`PC1`: 1000×1000×1000)** | Preset `FTG_2WAY` tạo `PC1-M1` (`D07`, `BarDir=X`, `DistDir=Y`, `HookL/R=150`, `@1=150`) và `PC1-M2`. Muốn thêm lớp trên `TOP.X/Y` hay đai quây giữa móng → chuột phải → **Chèn Thanh Chính**, đổi `Role=TOP.X`, `HookL/R=0` (hoặc giữ 150), số lượng & chiều dài nhảy ngay. |
| **Dầm / Giằng (`GBX1`: 48.6m × 0.6m × 1.0m)** | Preset `BEAM_FULL` tạo sẵn 7 thanh: `T1` (`SL.m=4`), `B1` (`SL.m=4`), `SK` (thép giá thành dầm `Part=2`, tự cộng neo `2*40D` ở `+-Lth`), `GG` (gia cường gối `X*0.25`, `Part=2`), `GN` (gia cường nhịp `X*0.7`), đai `ST1` (`PERIM_YH` + `ZONE2_X` `@150/@200`) và đai `ST2` (`TIE_Y`). Muốn tăng/giảm số cây chủ chỉ cần sửa trực tiếp `SL.m`; muốn tính cắt trừ cột dầm giao nhau gõ số âm vào `+-Pb` hoặc `+-Lth`. |
| **Cột nối sole (`C1`: 0.6 × 0.6 × 4.5m)** | Preset `COL_STAGGER` tự chia 2 nhóm: `C1A` (`SL.m=4`, `+-Lth = 45D + 600mm = 1,320mm` → `Mid = 5,750mm`) và `C1B` (`SL.m=4`, `+-Lth = 45D = 720mm` → `Mid = 5,150mm`), cùng đai bao `D1`, đai lồng `D2` (`PERIM_XY_INNER`), đai C `D3`. Đổi `SPEC_Col_StaggerOffset` bên `QS_Spec` từ `600` lên `800` → nhóm A tự dài thêm `200mm` ngay lập tức. |
| **Vách 2 mặt (`W1`: 6.0 × 0.3 × 4.0m)** | Preset `WALL_2FACE` tự tạo: `VD` (`Part=2`, `BarDir=H`, `DistDir=X @200` rải dọc chiều dài 6m), `HN` (`Part=2`, `BarDir=X`, `DistDir=H @200` rải dọc chiều cao 4m), `DC` (đai C `TIE_Y`, `DistDir=XH_GRID @600×@600`). |
| **Sàn 2 lớp & Chân chó (`GS1`)** | Preset `SLAB_2WAY_2LAYER` (hoặc `SLAB_WITH_HATS` nếu sàn dùng thép mũ) tự tạo đủ thép lớp dưới/trên + `CK` (chân chó `CHAIR_H`, lưới `XY_GRID @600×@600`). |

---

## 6. Cách chèn thanh tùy biến & Chọn 41 dáng thép

1. **Chèn thêm dòng thép mới trong Block**:
   - Đứng tại bất kỳ dòng nào bên trong Block cấu kiện cần thêm (kể cả dòng tiêu đề vàng `BLK_HDR` hay dòng cuối `--- End element ---`).
   - Bấm **Chuột phải** (hoặc Ribbon **`Bóc tách Bán Tự Động → Công cụ Thép CK`**) chọn 1 trong 3 nút:
     - **`Chèn Thanh Chính vào Block`** (`MAIN`)
     - **`Chèn Thanh Đai / C-Tie vào Block`** (`STIR`)
     - **`Chèn Thanh Tùy Biến (CUSTOM)`** (`CUSTOM`)
   - Add-in tự chèn dòng đúng vị trí, tự điền toàn bộ công thức Vùng A + Vùng C và cập nhật vùng `SUBTOTAL` của Block cấu kiện đó.
2. **Dùng Nút Chọn Dáng Thép (41 hình `D01..D41`)**:
   - Đứng ở dòng thép muốn đổi dáng → bấm **`Công cụ Thép CK → Chọn Dáng Thép (41 dáng L1..L7)`** (hoặc chuột phải → *Chọn Dáng Thép*).
   - Chọn mẫu trong 41 hình thép → bấm **Đổi dáng dòng đang chọn** (hoặc **Chèn tại dòng chọn**).
   - Nếu mẫu hình sử dụng từ 4 đoạn trở lên (vd `D17`, `D18`, `D22`...), các ô **`Le.1a`, `Le.1b`, `Ri.2b`, `Ri.2a`** tương ứng sẽ **tự động tô màu vàng nhạt** để nhắc bạn điền kích thước đoạn uốn phụ.

---

## 7. Mẹo & Lưu ý vận hành an toàn

- **Không gõ đè công thức ở Vùng A (`B..H`) và Vùng C (`Z..AL`, trừ 4 ô đoạn phụ `AA, AB, AD, AE`)**: Khi cần sửa kích thước, hãy sửa ở `QS_Data` (hoặc dùng 4 ô `SL.m`, `L.m`, `+-Pb`, `+-Lth` ở Vùng B).
- **Xóa bớt 1 thanh thép**: Có thể xóa hẳn dòng đó (nếu xóa hết dòng không cần thiết) hoặc đơn giản nhất là gõ `0` vào `SL.m` (`U`) để giữ lại vết kiểm toán mà khối lượng về `0.00 kg`.
- **Đồng bộ không mất dữ liệu**: Lệnh **`Nạp / Cập nhật Block Thép từ QS_Data`** luôn kiểm tra `ElementID` đã tồn tại hay chưa; các Block bạn đã tùy biến trước đó không bao giờ bị ghi đè. Khi cần tái sinh hoàn toàn 1 cấu kiện theo preset mới, chỉ cần xóa Block cũ của `ElementID` đó trên `QS_RebarConfig` rồi bấm Nạp lại.
- **Xuất hồ sơ gửi CĐT**: Luôn dùng nút **`Xuất .xlsx Sạch`** (hoặc gửi file đã chạy **`3. Sinh Detail + TKT`**): mọi công thức trên `QS_RebarConfig`, `TKT_*` và `*_Detail` đều là hàm Excel cơ bản 2007–365 (`IF`, `INDEX/MATCH`, `SUBTOTAL`, `ROUND`, `ROUNDUP`, `PI`), mở trên máy không cài Add-in không bao giờ lỗi `#NAME?`.
