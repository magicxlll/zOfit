using System.Globalization;
using ZOfit.Engine.Models;

namespace ZOfit.Engine.Generators;

public sealed class InfraEngine
{
    public GeneratedSheetDto Generate(GenerateInfraRequest req)
    {
        var rows = new List<SheetCellRowDto>();
        rows.Add(RcDetailGenerator.MakeRow("TITLE", 0, "HDR_MARKER", "HDR", "ZOFIT_INFRA_V2", "INFRASTRUCTURE DETAIL TAKEOFF / BÓC TÁCH HẠ TẦNG KỸ THUẬT", req.Spec.ProjectName, "", "", "", "", "", "", "", "", "", "", ""));
        rows.Add(RcDetailGenerator.MakeRow("BANNER", 0, "HDR_PROJ", "HDR", "", "Project:", $"=IFERROR(CFG_ProjectName,\"{req.Spec.ProjectName.Replace("\"", "\"\"")}\")", "", "Spec:", "TCVN 7957 / 9113 / 10380 & Asphalt Road", "", "", "", "", "", "", "HDR_PROJ", "HDR"));
        rows.Add(RcDetailGenerator.MakeRow("BLANK", 0, "", "HDR"));
        rows.Add(RcDetailGenerator.MakeRow("WORK_HDR", 0, "HDR_COLS", "HDR", "Code", "Description", "Spec / Dim", "Unit", "NO", "L (m)", "W (m)", "H / t (m)", "a", "b", "c", "Total Q'TY", "TagID (Sync)", "SectionKey"));
        rows.Add(RcDetailGenerator.MakeRow("BLANK", 0, "", "HDR"));

        var manualMap = GroupManual(req.ExistingRows);
        var items = req.Items ?? new List<InfraItemDto>();

        BuildPipelineSection(rows, items, manualMap);
        BuildRoadSection(rows, items, manualMap);
        BuildFenceSection(rows, items, manualMap);

        return new GeneratedSheetDto
        {
            SheetName = "Civil_Detail",
            Building = "CIVIL",
            AddedCount = items.Count,
            UpdatedCount = 0,
            PreservedManualCount = manualMap.Values.Sum(v => v.Count),
            RemovedFlaggedCount = 0,
            Rows = rows
        };
    }

    private void BuildPipelineSection(List<SheetCellRowDto> rows, List<InfraItemDto> items, Dictionary<string, List<ExistingDetailRowDto>> manualMap)
    {
        var pipes = items.Where(i => IsCat(i, "PIPE", "CULVERT")).ToList();
        if (pipes.Count == 0 && !manualMap.ContainsKey("INF_PIPE")) return;

        rows.Add(RcDetailGenerator.MakeRow("WORK_HDR", 0, "SEC_PIPE", "INF_PIPE", "I", "DRAINAGE PIPELINE & BOX CULVERT WORKS / TUYẾN CỐNG & MƯƠNG THOÁT NƯỚC", "", "", null, null, null, null, null, null, null, null, "SEC_PIPE", "INF_PIPE"));

        // 1. Supply & Installation Length
        int hdrLen = NextRow(rows);
        rows.Add(RcDetailGenerator.MakeRow("ITEM_HDR", 1, "INF_PIPE_LEN", "INF_PIPE", "1", "Supply & install drainage pipe / culvert / Lắp đặt tuyến ống, cống", "", "m", null, null, null, null, null, null, null, 0, "INF_PIPE_LEN", "INF_PIPE"));
        int s1 = NextRow(rows);
        foreach (var p in pipes)
        {
            int r = NextRow(rows);
            int d = p.DataRow;
            rows.Add(RcDetailGenerator.MakeRow("DATA", 2, $"{p.InfraId}::LEN", "INF_PIPE",
                p.Code, p.Description, $"D{(int)p.DiaMm} / W{Fmt(p.WidthM)}", "m",
                $"=QS_Infra!E{d}", $"=QS_Infra!F{d}", null, null, null, null, null, $"=PRODUCT(E{r}:K{r})", $"{p.InfraId}::LEN", "INF_PIPE"));
        }
        int e1 = NextRow(rows) - 1;
        if (e1 >= s1) rows[hdrLen - 1].Values[11] = $"=SUBTOTAL(9,L{s1}:L{e1})";

        // 2. Trench Excavation with Talus slope
        int hdrExc = NextRow(rows);
        rows.Add(RcDetailGenerator.MakeRow("ITEM_HDR", 1, "INF_PIPE_EXC", "INF_PIPE", "2", "Trench excavation with talus slope (SPEC_Pipe_Slope) / Đào rãnh cống có taluy", "", "m3", null, null, null, null, null, null, null, 0, "INF_PIPE_EXC", "INF_PIPE"));
        int s2 = NextRow(rows);
        foreach (var p in pipes)
        {
            int r = NextRow(rows);
            int d = p.DataRow;
            string dOutExpr = ResolveOuterWidthExpr(p, d);
            // B_bot = D_out + 0.6; Trapezoid: (B_bot + B_bot + 2*m*H)/2 * H * L = (B_bot + m*H) * H * L
            rows.Add(RcDetailGenerator.MakeRow("DATA", 2, $"{p.InfraId}::EXC", "INF_PIPE",
                p.Code, "Excavate " + p.Code, "m=SPEC_Pipe_Slope", "m3",
                $"=QS_Infra!E{d}", $"=QS_Infra!F{d}", $"={dOutExpr}+0.6", $"=QS_Infra!H{d}", null, null, null,
                $"=E{r}*F{r}*(G{r}+SPEC_Pipe_Slope*H{r})*H{r}", $"{p.InfraId}::EXC", "INF_PIPE"));
        }
        int e2 = NextRow(rows) - 1;
        if (e2 >= s2) rows[hdrExc - 1].Values[11] = $"=SUBTOTAL(9,L{s2}:L{e2})";

        // 3. Bedding Cushion
        int hdrBed = NextRow(rows);
        rows.Add(RcDetailGenerator.MakeRow("ITEM_HDR", 1, "INF_PIPE_BED", "INF_PIPE", "3", "Sand / crushed stone bedding under pipe / Đệm cát đá đáy rãnh cống", "", "m3", null, null, null, null, null, null, null, 0, "INF_PIPE_BED", "INF_PIPE"));
        int s3 = NextRow(rows);
        foreach (var p in pipes)
        {
            int r = NextRow(rows);
            int d = p.DataRow;
            string dOutExpr = ResolveOuterWidthExpr(p, d);
            rows.Add(RcDetailGenerator.MakeRow("DATA", 2, $"{p.InfraId}::BED", "INF_PIPE",
                p.Code, "Bedding " + p.Code, "", "m3",
                $"=QS_Infra!E{d}", $"=QS_Infra!F{d}", $"={dOutExpr}+0.6", $"=IF(QS_Infra!L{d}>0,QS_Infra!L{d},SPEC_Pipe_Bedding)", null, null, null,
                $"=PRODUCT(E{r}:K{r})", $"{p.InfraId}::BED", "INF_PIPE"));
        }
        int e3 = NextRow(rows) - 1;
        if (e3 >= s3) rows[hdrBed - 1].Values[11] = $"=SUBTOTAL(9,L{s3}:L{e3})";

        // 4. Trench Backfill (Excavation - Bedding - Pipe Occupied Volume)
        int hdrBk = NextRow(rows);
        rows.Add(RcDetailGenerator.MakeRow("ITEM_HDR", 1, "INF_PIPE_BK", "INF_PIPE", "4", "Trench backfill (deduct bedding & pipe displacement) / Đắp trả rãnh cống", "", "m3", null, null, null, null, null, null, null, 0, "INF_PIPE_BK", "INF_PIPE"));
        int s4 = NextRow(rows);
        for (int i = 0; i < pipes.Count; i++)
        {
            var p = pipes[i];
            int d = p.DataRow;
            int rExc = s2 + i;
            int rBed = s3 + i;
            string dOutExpr = ResolveOuterWidthExpr(p, d);
            string occupiedFormula = IsCat(p, "PIPE")
                ? $"(PI()*(({dOutExpr})/2)^2*QS_Infra!F{d}*QS_Infra!E{d})"
                : $"(({dOutExpr})*({dOutExpr})*QS_Infra!F{d}*QS_Infra!E{d})";
            rows.Add(RcDetailGenerator.MakeRow("DATA", 2, $"{p.InfraId}::BK", "INF_PIPE",
                p.Code, "Backfill " + p.Code, "Exc - Bed - Pipe", "m3",
                null, null, null, null, null, null, null,
                $"=MAX(0,L{rExc}-L{rBed}-{occupiedFormula})", $"{p.InfraId}::BK", "INF_PIPE"));
        }
        int e4 = NextRow(rows) - 1;
        if (e4 >= s4) rows[hdrBk - 1].Values[11] = $"=SUBTOTAL(9,L{s4}:L{e4})";

        // 5. Geotextile wrap at joints
        int hdrGeo = NextRow(rows);
        rows.Add(RcDetailGenerator.MakeRow("ITEM_HDR", 1, "INF_PIPE_GEO", "INF_PIPE", "5", "Geotextile joint wrap (2.5m/seg, 0.4m width) / Vải địa kỹ thuật bọc mối nối", "", "m2", null, null, null, null, null, null, null, 0, "INF_PIPE_GEO", "INF_PIPE"));
        int s5 = NextRow(rows);
        foreach (var p in pipes)
        {
            int r = NextRow(rows);
            int d = p.DataRow;
            string dOutExpr = ResolveOuterWidthExpr(p, d);
            rows.Add(RcDetailGenerator.MakeRow("DATA", 2, $"{p.InfraId}::GEO", "INF_PIPE",
                p.Code, "Geotextile " + p.Code, "", "m2",
                $"=ROUNDUP(QS_Infra!F{d}/2.5,0)", $"=PI()*({dOutExpr})", 0.4, null, null, null, null,
                $"=PRODUCT(E{r}:K{r})", $"{p.InfraId}::GEO", "INF_PIPE"));
        }
        int e5 = NextRow(rows) - 1;
        if (e5 >= s5) rows[hdrGeo - 1].Values[11] = $"=SUBTOTAL(9,L{s5}:L{e5})";

        AppendManual(rows, manualMap, "INF_PIPE");
        rows.Add(RcDetailGenerator.MakeRow("BLANK", 0, "", "INF_PIPE"));
    }

    private void BuildRoadSection(List<SheetCellRowDto> rows, List<InfraItemDto> items, Dictionary<string, List<ExistingDetailRowDto>> manualMap)
    {
        var roads = items.Where(i => IsCat(i, "ROAD_BTXM", "ROAD_ASPHALT")).ToList();
        if (roads.Count == 0 && !manualMap.ContainsKey("INF_ROAD")) return;

        rows.Add(RcDetailGenerator.MakeRow("WORK_HDR", 0, "SEC_ROAD", "INF_ROAD", "II", "INTERNAL ROADS, DOWELS, CURBS & SIDEWALK / ĐƯỜNG NỘI BỘ, BÓ VỈA & VỈA HÈ", "", "", null, null, null, null, null, null, null, null, "SEC_ROAD", "INF_ROAD"));

        // 1. Pavement Area & Base Layers
        int hdrArea = NextRow(rows);
        rows.Add(RcDetailGenerator.MakeRow("ITEM_HDR", 1, "INF_ROAD_AREA", "INF_ROAD", "1", "Road pavement surface & tack coat / Diện tích mặt đường hoàn thiện", "", "m2", null, null, null, null, null, null, null, 0, "INF_ROAD_AREA", "INF_ROAD"));
        int rs1 = NextRow(rows);
        foreach (var rd in roads)
        {
            int r = NextRow(rows);
            int d = rd.DataRow;
            rows.Add(RcDetailGenerator.MakeRow("DATA", 2, $"{rd.InfraId}::AREA", "INF_ROAD",
                rd.Code, rd.Description, rd.Category, "m2",
                $"=QS_Infra!E{d}", $"=QS_Infra!F{d}", $"=QS_Infra!G{d}", null, null, null, null,
                $"=PRODUCT(E{r}:K{r})", $"{rd.InfraId}::AREA", "INF_ROAD"));
        }
        int re1 = NextRow(rows) - 1;
        if (re1 >= rs1) rows[hdrArea - 1].Values[11] = $"=SUBTOTAL(9,L{rs1}:L{re1})";

        // 2. Concrete Road Volume + Subbase CPDD1/2 + Nylon + Dowels
        var btxmRoads = roads.Where(x => IsCat(x, "ROAD_BTXM")).ToList();
        if (btxmRoads.Count > 0)
        {
            int hdrConc = NextRow(rows);
            rows.Add(RcDetailGenerator.MakeRow("ITEM_HDR", 1, "INF_ROAD_CONC", "INF_ROAD", "2", "Concrete pavement M300 / Bê tông mặt đường M300", "", "m3", null, null, null, null, null, null, null, 0, "INF_ROAD_CONC", "INF_ROAD"));
            int cs = NextRow(rows);
            foreach (var rd in btxmRoads)
            {
                int r = NextRow(rows);
                int d = rd.DataRow;
                rows.Add(RcDetailGenerator.MakeRow("DATA", 2, $"{rd.InfraId}::CONC", "INF_ROAD",
                    rd.Code, "Concrete " + rd.Code, "", "m3",
                    $"=QS_Infra!E{d}", $"=QS_Infra!F{d}", $"=QS_Infra!G{d}", $"=QS_Infra!H{d}", null, null, null,
                    $"=PRODUCT(E{r}:K{r})", $"{rd.InfraId}::CONC", "INF_ROAD"));
            }
            int ce = NextRow(rows) - 1;
            rows[hdrConc - 1].Values[11] = $"=SUBTOTAL(9,L{cs}:L{ce})";

            int hdrDow = NextRow(rows);
            rows.Add(RcDetailGenerator.MakeRow("ITEM_HDR", 1, "INF_ROAD_DOWEL", "INF_ROAD", "3", "Dowel steel D20 at expansion joints / Thép truyền lực D20 khe co giãn", "SPEC_Road_DowelLen", "kg", null, null, null, null, null, null, null, 0, "INF_ROAD_DOWEL", "INF_ROAD"));
            int ds = NextRow(rows);
            foreach (var rd in btxmRoads)
            {
                int r = NextRow(rows);
                int d = rd.DataRow;
                rows.Add(RcDetailGenerator.MakeRow("DATA", 2, $"{rd.InfraId}::DOWEL", "INF_ROAD",
                    rd.Code, "Dowels D20 " + rd.Code, "D20 @300", "kg",
                    $"=INT(QS_Infra!F{d}/IF(QS_Infra!O{d}>0,QS_Infra!O{d},5))*(INT(QS_Infra!G{d}/SPEC_Road_DowelSpace)+1)",
                    "=SPEC_Road_DowelLen", null, null, "=ROUND(PI()*(20/1000)^2/4*7850,3)", null, null,
                    $"=PRODUCT(E{r}:K{r})", $"{rd.InfraId}::DOWEL", "INF_ROAD"));
            }
            int de = NextRow(rows) - 1;
            rows[hdrDow - 1].Values[11] = $"=SUBTOTAL(9,L{ds}:L{de})";
        }

        // 3. CPDD Subbase Type 1 & Type 2 (widened +0.2m each side => +0.4m)
        int hdrBase1 = NextRow(rows);
        rows.Add(RcDetailGenerator.MakeRow("ITEM_HDR", 1, "INF_ROAD_CPDD1", "INF_ROAD", "4", "Crushed stone subbase Type 1 (+0.2m each side) / Cấp phối đá dăm loại 1", "", "m3", null, null, null, null, null, null, null, 0, "INF_ROAD_CPDD1", "INF_ROAD"));
        int b1s = NextRow(rows);
        foreach (var rd in roads)
        {
            int r = NextRow(rows);
            int d = rd.DataRow;
            rows.Add(RcDetailGenerator.MakeRow("DATA", 2, $"{rd.InfraId}::CPDD1", "INF_ROAD",
                rd.Code, "CPDD1 " + rd.Code, "", "m3",
                $"=QS_Infra!E{d}", $"=QS_Infra!F{d}", $"=QS_Infra!G{d}+0.4", $"=QS_Infra!J{d}", null, null, null,
                $"=PRODUCT(E{r}:K{r})", $"{rd.InfraId}::CPDD1", "INF_ROAD"));
        }
        int b1e = NextRow(rows) - 1;
        rows[hdrBase1 - 1].Values[11] = $"=SUBTOTAL(9,L{b1s}:L{b1e})";

        // 4. Sidewalk & Curbs
        var swRoads = roads.Where(x => x.ExtraSidewalkWM > 0 || x.CurbLenM > 0).ToList();
        if (swRoads.Count > 0)
        {
            int hdrCurb = NextRow(rows);
            rows.Add(RcDetailGenerator.MakeRow("ITEM_HDR", 1, "INF_CURB", "INF_ROAD", "5", "Concrete curbs (180x300) / Bó vỉa bê tông 180x300", "0.18x0.30", "m3", null, null, null, null, null, null, null, 0, "INF_CURB", "INF_ROAD"));
            int cbs = NextRow(rows);
            foreach (var rd in swRoads)
            {
                int r = NextRow(rows);
                int d = rd.DataRow;
                rows.Add(RcDetailGenerator.MakeRow("DATA", 2, $"{rd.InfraId}::CURB", "INF_ROAD",
                    rd.Code, "Curb " + rd.Code, "", "m3",
                    $"=QS_Infra!E{d}", $"=QS_Infra!N{d}", 0.18, 0.30, null, null, null,
                    $"=PRODUCT(E{r}:K{r})", $"{rd.InfraId}::CURB", "INF_ROAD"));
            }
            int cbe = NextRow(rows) - 1;
            rows[hdrCurb - 1].Values[11] = $"=SUBTOTAL(9,L{cbs}:L{cbe})";

            int hdrSw = NextRow(rows);
            rows.Add(RcDetailGenerator.MakeRow("ITEM_HDR", 1, "INF_SIDEWALK", "INF_ROAD", "6", "Sidewalk Terrazzo tiles / Lát gạch Terrazzo vỉa hè (2 sides)", "", "m2", null, null, null, null, null, null, null, 0, "INF_SIDEWALK", "INF_ROAD"));
            int sws = NextRow(rows);
            foreach (var rd in swRoads.Where(x => x.ExtraSidewalkWM > 0))
            {
                int r = NextRow(rows);
                int d = rd.DataRow;
                rows.Add(RcDetailGenerator.MakeRow("DATA", 2, $"{rd.InfraId}::SW", "INF_ROAD",
                    rd.Code, "Sidewalk " + rd.Code, "", "m2",
                    2, $"=QS_Infra!F{d}", $"=QS_Infra!M{d}", null, null, null, null,
                    $"=PRODUCT(E{r}:K{r})", $"{rd.InfraId}::SW", "INF_ROAD"));
            }
            int swe = NextRow(rows) - 1;
            if (swe >= sws) rows[hdrSw - 1].Values[11] = $"=SUBTOTAL(9,L{sws}:L{swe})";
        }

        AppendManual(rows, manualMap, "INF_ROAD");
        rows.Add(RcDetailGenerator.MakeRow("BLANK", 0, "", "INF_ROAD"));
    }

    private void BuildFenceSection(List<SheetCellRowDto> rows, List<InfraItemDto> items, Dictionary<string, List<ExistingDetailRowDto>> manualMap)
    {
        var fences = items.Where(i => IsCat(i, "FENCE")).ToList();
        if (fences.Count == 0 && !manualMap.ContainsKey("INF_FENCE")) return;

        rows.Add(RcDetailGenerator.MakeRow("WORK_HDR", 0, "SEC_FENCE", "INF_FENCE", "III", "PERIMETER FENCE WORKS / TƯỜNG RÀO KHUÔN VIÊN", "", "", null, null, null, null, null, null, null, null, "SEC_FENCE", "INF_FENCE"));

        // 1. Plinth beam Concrete & Columns Concrete
        int hdrCol = NextRow(rows);
        rows.Add(RcDetailGenerator.MakeRow("ITEM_HDR", 1, "INF_FENCE_CONC", "INF_FENCE", "1", "Fence plinth beam & RC columns (250x250) / Giằng móng & Cột BTCT hàng rào", "", "m3", null, null, null, null, null, null, null, 0, "INF_FENCE_CONC", "INF_FENCE"));
        int s = NextRow(rows);
        foreach (var f in fences)
        {
            int rBeam = NextRow(rows);
            int d = f.DataRow;
            rows.Add(RcDetailGenerator.MakeRow("DATA", 2, $"{f.InfraId}::PLINTH", "INF_FENCE",
                f.Code, "Plinth beam " + f.Code, "200x300", "m3",
                1, $"=QS_Infra!F{d}", 0.2, 0.3, null, null, null, $"=PRODUCT(E{rBeam}:K{rBeam})", $"{f.InfraId}::PLINTH", "INF_FENCE"));

            int rCol = NextRow(rows);
            rows.Add(RcDetailGenerator.MakeRow("DATA", 2, $"{f.InfraId}::FCOL", "INF_FENCE",
                f.Code, "Fence RC cols " + f.Code, "250x250", "m3",
                $"=INT(QS_Infra!F{d}/IF(QS_Infra!P{d}>0,QS_Infra!P{d},3))+1", 0.25, 0.25, $"=QS_Infra!H{d}", null, null, null,
                $"=PRODUCT(E{rCol}:K{rCol})", $"{f.InfraId}::FCOL", "INF_FENCE"));
        }
        int e = NextRow(rows) - 1;
        rows[hdrCol - 1].Values[11] = $"=SUBTOTAL(9,L{s}:L{e})";

        // 2. Fence Brick Wall (deduct columns) & Steel Fence Panel above
        int hdrBr = NextRow(rows);
        rows.Add(RcDetailGenerator.MakeRow("ITEM_HDR", 1, "INF_FENCE_BRICK", "INF_FENCE", "2", "Fence brick wall (deduct cols) / Tường xây gạch hàng rào", "t=200mm", "m3", null, null, null, null, null, null, null, 0, "INF_FENCE_BRICK", "INF_FENCE"));
        int bs = NextRow(rows);
        foreach (var f in fences)
        {
            int r = NextRow(rows);
            int d = f.DataRow;
            rows.Add(RcDetailGenerator.MakeRow("DATA", 2, $"{f.InfraId}::BRICK", "INF_FENCE",
                f.Code, "Brick wall " + f.Code, "", "m3",
                1, $"=QS_Infra!F{d}-(INT(QS_Infra!F{d}/IF(QS_Infra!P{d}>0,QS_Infra!P{d},3))+1)*0.25", 0.2, $"=IF(QS_Infra!O{d}>0,QS_Infra!O{d},0.8)", null, null, null,
                $"=PRODUCT(E{r}:K{r})", $"{f.InfraId}::BRICK", "INF_FENCE"));
        }
        int be = NextRow(rows) - 1;
        rows[hdrBr - 1].Values[11] = $"=SUBTOTAL(9,L{bs}:L{be})";

        int hdrSt = NextRow(rows);
        rows.Add(RcDetailGenerator.MakeRow("ITEM_HDR", 1, "INF_FENCE_STEEL", "INF_FENCE", "3", "Steel grille / wire fence panel / Khung hàng rào thép hộp", "", "m2", null, null, null, null, null, null, null, 0, "INF_FENCE_STEEL", "INF_FENCE"));
        int sts = NextRow(rows);
        foreach (var f in fences)
        {
            int r = NextRow(rows);
            int d = f.DataRow;
            rows.Add(RcDetailGenerator.MakeRow("DATA", 2, $"{f.InfraId}::STEEL", "INF_FENCE",
                f.Code, "Steel panel " + f.Code, "", "m2",
                1, $"=QS_Infra!F{d}-(INT(QS_Infra!F{d}/IF(QS_Infra!P{d}>0,QS_Infra!P{d},3))+1)*0.25", $"=MAX(0,QS_Infra!H{d}-IF(QS_Infra!O{d}>0,QS_Infra!O{d},0.8))", null, null, null, null,
                $"=PRODUCT(E{r}:K{r})", $"{f.InfraId}::STEEL", "INF_FENCE"));
        }
        int ste = NextRow(rows) - 1;
        rows[hdrSt - 1].Values[11] = $"=SUBTOTAL(9,L{sts}:L{ste})";

        AppendManual(rows, manualMap, "INF_FENCE");
    }

    private static string ResolveOuterWidthExpr(InfraItemDto p, int d)
    {
        if (IsCat(p, "PIPE"))
        {
            // TCVN 9113 wall thickness: <=400 -> +100mm; <=800 -> +160mm; >800 -> +200mm
            return $"(QS_Infra!I{d}+IF(QS_Infra!I{d}<=400,100,IF(QS_Infra!I{d}<=800,160,200)))/1000";
        }
        return $"QS_Infra!G{d}";
    }

    private static bool IsCat(InfraItemDto item, params string[] cats) =>
        cats.Any(c => string.Equals(item.Category?.Trim(), c, StringComparison.OrdinalIgnoreCase));

    private static string Fmt(double v) => v.ToString("0.###", CultureInfo.InvariantCulture);

    private static Dictionary<string, List<ExistingDetailRowDto>> GroupManual(IReadOnlyList<ExistingDetailRowDto>? existing)
    {
        var dict = new Dictionary<string, List<ExistingDetailRowDto>>(StringComparer.OrdinalIgnoreCase);
        if (existing == null) return dict;
        foreach (var row in existing)
        {
            if (!string.Equals(row.TagId, "MANUAL", StringComparison.OrdinalIgnoreCase)) continue;
            string sec = string.IsNullOrWhiteSpace(row.SectionKey) ? "INF_PIPE" : row.SectionKey.Trim();
            if (!dict.TryGetValue(sec, out var list))
            {
                list = new List<ExistingDetailRowDto>();
                dict[sec] = list;
            }
            list.Add(row);
        }
        return dict;
    }

    private static void AppendManual(List<SheetCellRowDto> rows, Dictionary<string, List<ExistingDetailRowDto>> map, string sec)
    {
        if (!map.TryGetValue(sec, out var list)) return;
        foreach (var m in list)
        {
            int r = NextRow(rows);
            rows.Add(RcDetailGenerator.MakeRow("MANUAL", 2, "MANUAL", sec,
                m.Code, m.Description, m.SpecText, m.Unit,
                string.IsNullOrWhiteSpace(m.NoExpr) ? null : m.NoExpr,
                string.IsNullOrWhiteSpace(m.LExpr) ? null : m.LExpr,
                string.IsNullOrWhiteSpace(m.WExpr) ? null : m.WExpr,
                string.IsNullOrWhiteSpace(m.HExpr) ? null : m.HExpr,
                string.IsNullOrWhiteSpace(m.AExpr) ? null : m.AExpr,
                string.IsNullOrWhiteSpace(m.BExpr) ? null : m.BExpr,
                string.IsNullOrWhiteSpace(m.CExpr) ? null : m.CExpr,
                string.IsNullOrWhiteSpace(m.QtyExpr) ? $"=PRODUCT(E{r}:K{r})" : m.QtyExpr,
                "MANUAL", sec));
        }
    }

    private static int NextRow(List<SheetCellRowDto> rows) => rows.Count + 1;
}
