using System.Globalization;
using ZOfit.Engine.Models;

namespace ZOfit.Engine.Generators;

public sealed class RcDetailGenerator
{
    public const int TotalCols = 14; // Columns A..N

    public GeneratedSheetDto Generate(string building, ProjectSpecDto spec, IReadOnlyList<ElementDataDto> elements, IReadOnlyList<ExistingDetailRowDto>? existingRows = null)
    {
        var bElements = elements
            .Where(e => string.Equals(e.Building, building, StringComparison.OrdinalIgnoreCase))
            .ToList();

        var rows = new List<SheetCellRowDto>();
        string cleanSheet = SanitizeSheetName($"{building}_Detail");

        // Header Rows 1..5
        rows.Add(MakeRow("TITLE", 0, "HDR_MARKER", "HDR", "ZOFIT_DETAIL_V2", $"{building.ToUpperInvariant()} - DETAIL QUANTITY TAKEOFF (STRUCTURAL & SOIL)", spec.ProjectName, "", "", "", "", "", "", "", "", "", "", ""));
        rows.Add(MakeRow("BANNER", 0, "HDR_PROJ", "HDR", "", "Project:", $"=IFERROR(CFG_ProjectName,\"{EscapeExcel(spec.ProjectName)}\")", "", "Preset:", "=IFERROR(CFG_Preset,\"FDI / TCVN\")", "", "", "", "", "", "", "HDR_PROJ", "HDR"));
        rows.Add(MakeRow("BLANK", 0, "", "HDR"));
        rows.Add(MakeRow("WORK_HDR", 0, "HDR_COLS", "HDR", "Code", "Description", "Spec / Dim", "Unit", "NO", "L (m)", "W (m)", "H / t (m)", "a", "b", "c", "Total Q'TY", "TagID (Sync)", "SectionKey"));
        rows.Add(MakeRow("BLANK", 0, "", "HDR"));

        var manualBySection = GroupManualRows(existingRows);

        // SECTION 1: PILING WORKS
        var piles = bElements.Where(e => IsGroup(e, "Pile")).ToList();
        if (piles.Count > 0 || manualBySection.ContainsKey("PILE"))
        {
            BuildPileSection(rows, piles, spec, manualBySection);
        }

        // SECTION 2: SOIL WORKS
        var buriedFtg = bElements.Where(e => IsBuried(e) && IsGroup(e, "Foundation", "Tank", "Manhole")).ToList();
        var buriedGb = bElements.Where(e => IsBuried(e) && IsGroup(e, "GroundBeam")).ToList();
        var sogList = bElements.Where(e => IsGroup(e, "SlabOnGrade")).ToList();
        if (buriedFtg.Count > 0 || buriedGb.Count > 0 || sogList.Count > 0 || manualBySection.ContainsKey("SOIL_EXC") || manualBySection.ContainsKey("SOIL_BACKFILL"))
        {
            BuildSoilSection(rows, buriedFtg, buriedGb, sogList, bElements, manualBySection);
        }

        // SECTION 3: FORMWORK WORKS
        BuildFormworkSection(rows, bElements, manualBySection);

        // SECTION 4: CONCRETE WORKS
        BuildConcreteSection(rows, bElements, manualBySection);

        // Sync any flagged removed items? Notice: Fresh build produces exact structure;
        // SyncEngine can adjust counts & preserve extra manual rows anywhere.
        int manualPreserved = manualBySection.Values.Sum(v => v.Count);

        return new GeneratedSheetDto
        {
            SheetName = cleanSheet,
            Building = building,
            AddedCount = bElements.Count,
            UpdatedCount = 0,
            PreservedManualCount = manualPreserved,
            RemovedFlaggedCount = 0,
            Rows = rows
        };
    }

    private void BuildPileSection(List<SheetCellRowDto> rows, List<ElementDataDto> piles, ProjectSpecDto spec, Dictionary<string, List<ExistingDetailRowDto>> manualMap)
    {
        rows.Add(MakeRow("WORK_HDR", 0, "SEC_PILE", "PILE", "I", "PILE WORKS / CÔNG TÁC CỌC PHC", "", "", null, null, null, null, null, null, null, null, "SEC_PILE", "PILE"));

        // 1. Supply & Pressing
        int hdrSupply = NextRow(rows);
        rows.Add(MakeRow("ITEM_HDR", 1, "PILE_SUPPLY_HDR", "PILE", "1", "Supply & pressing PHC piles / Cung cấp & ép cọc PHC", "", "m", null, null, null, null, null, null, null, 0, "PILE_SUPPLY_HDR", "PILE"));
        int s1 = NextRow(rows);
        foreach (var p in piles)
        {
            int r = NextRow(rows);
            int d = p.DataRow;
            rows.Add(MakeRow("DATA", 2, $"{p.ElementId}::PILE_LEN", "PILE",
                p.Mark, p.Description, $"D{Fmt(p.WidthM > 0 ? p.WidthM * 1000 : 400)}", "m",
                $"=QS_Data!F{d}", $"=QS_Data!G{d}", null, null, null, null, null, $"=PRODUCT(E{r}:K{r})", $"{p.ElementId}::PILE_LEN", "PILE"));
        }
        AppendManualRows(rows, manualMap, "PILE");
        int e1 = NextRow(rows) - 1;
        rows[hdrSupply - 1].Values[11] = RangeSubtotal(s1, e1);

        // 2. Cutting Pile Heads
        int hdrCut = NextRow(rows);
        rows.Add(MakeRow("ITEM_HDR", 1, "PILE_CUT_HDR", "PILE_CUT", "2", "Cut pile heads / Đập, cắt đầu cọc", "", "ea", null, null, null, null, null, null, null, 0, "PILE_CUT_HDR", "PILE_CUT"));
        int s2 = NextRow(rows);
        foreach (var p in piles)
        {
            int r = NextRow(rows);
            int d = p.DataRow;
            rows.Add(MakeRow("DATA", 2, $"{p.ElementId}::PILE_CUT", "PILE_CUT",
                p.Mark, "Cut head " + p.Mark, $"Cut={Fmt(spec.PileCutHeadM)}m", "ea",
                $"=QS_Data!F{d}", null, null, null, null, null, null, $"=PRODUCT(E{r}:K{r})", $"{p.ElementId}::PILE_CUT", "PILE_CUT"));
        }
        int e2 = NextRow(rows) - 1;
        rows[hdrCut - 1].Values[11] = RangeSubtotal(s2, e2);

        // 3. Welded Joints
        int hdrJnt = NextRow(rows);
        rows.Add(MakeRow("ITEM_HDR", 1, "PILE_JOINT_HDR", "PILE_JNT", "3", "Welded pile joints / Mối nối cọc hàn", "", "ea", null, null, null, null, null, null, null, 0, "PILE_JOINT_HDR", "PILE_JNT"));
        int s3 = NextRow(rows);
        foreach (var p in piles)
        {
            int r = NextRow(rows);
            int d = p.DataRow;
            rows.Add(MakeRow("DATA", 2, $"{p.ElementId}::PILE_JNT", "PILE_JNT",
                p.Mark, "Joints " + p.Mark, "", "ea",
                $"=QS_Data!F{d}", $"=MAX(0,ROUNDUP(QS_Data!G{d}/SPEC_Pile_SegLen,0)-1)", null, null, null, null, null, $"=PRODUCT(E{r}:K{r})", $"{p.ElementId}::PILE_JNT", "PILE_JNT"));
        }
        int e3 = NextRow(rows) - 1;
        rows[hdrJnt - 1].Values[11] = RangeSubtotal(s3, e3);

        // 4. Negative pressing & Static Load Test
        if (piles.Count > 0)
        {
            int rNeg = NextRow(rows);
            int firstD = piles[0].DataRow;
            rows.Add(MakeRow("ITEM_HDR", 1, "PILE_NEG_HDR", "PILE_NEG", "4", "Negative pressing piles / Ép cọc âm", "", "m",
                $"=L{hdrCut}", $"=MAX(0,IF(QS_Data!J{firstD}=0,SPEC_Elev_Natural,QS_Data!J{firstD})-QS_Data!K{firstD})", null, null, null, null, null, $"=PRODUCT(E{rNeg}:K{rNeg})", "PILE_NEG_HDR", "PILE_NEG"));

            int rTest = NextRow(rows);
            rows.Add(MakeRow("ITEM_HDR", 1, "PILE_TEST_HDR", "PILE_TEST", "5", "Static loading test (1% min 2) / Thí nghiệm nén tĩnh", "", "ea",
                $"=MAX(2,ROUNDUP(L{hdrCut}*0.01,0))", null, null, null, null, null, null, $"=PRODUCT(E{rTest}:K{rTest})", "PILE_TEST_HDR", "PILE_TEST"));
        }
        rows.Add(MakeRow("BLANK", 0, "", "PILE"));
    }

    private void BuildSoilSection(
        List<SheetCellRowDto> rows,
        List<ElementDataDto> buriedFtg,
        List<ElementDataDto> buriedGb,
        List<ElementDataDto> sogList,
        List<ElementDataDto> allElements,
        Dictionary<string, List<ExistingDetailRowDto>> manualMap)
    {
        rows.Add(MakeRow("WORK_HDR", 0, "SEC_SOIL", "SOIL", "II", "SOIL WORKS / CÔNG TÁC ĐẤT & NỀN", "", "", null, null, null, null, null, null, null, null, "SEC_SOIL", "SOIL"));

        // Track concrete structural rows to reuse later or keep direct calculation
        int excTotalRow = NextRow(rows);
        rows.Add(MakeRow("ITEM_HDR", 1, "SOIL_EXC_GROSS", "SOIL_EXC", "1", "Total Pit Excavation (Gross) / Tổng khối lượng đào hố móng", "", "m3", null, null, null, null, null, null, null, 0, "SOIL_EXC_GROSS", "SOIL_EXC"));

        // Topsoil stripping from SlabOnGrade
        int excSubStart = NextRow(rows);
        if (sogList.Count > 0)
        {
            rows.Add(MakeRow("SUB_HDR", 1, "SOIL_TOPSOIL_HDR", "SOIL_EXC", "", "Topsoil / Unsuitable soil stripping (SPEC_Topsoil_Thk)", "", "m3", null, null, null, null, null, null, null, null, "SOIL_TOPSOIL_HDR", "SOIL_EXC"));
            foreach (var s in sogList)
            {
                int r = NextRow(rows);
                int d = s.DataRow;
                rows.Add(MakeRow("DATA", 2, $"{s.ElementId}::TOPSOIL", "SOIL_EXC",
                    s.Mark, "Strip unsuitable soil - " + s.Mark, "", "m3",
                    $"=QS_Data!F{d}", $"=QS_Data!G{d}", $"=QS_Data!H{d}", "=SPEC_Topsoil_Thk", "=SPEC_Exc_SlopeK", null, null, $"=PRODUCT(E{r}:K{r})", $"{s.ElementId}::TOPSOIL", "SOIL_EXC"));
            }
        }

        // Foundation / Tank / Manhole pit excavation
        if (buriedFtg.Count > 0)
        {
            rows.Add(MakeRow("SUB_HDR", 1, "SOIL_FTG_HDR", "SOIL_EXC", "", "Foundation / Pit excavation / Đào đất hố móng", "", "m3", null, null, null, null, null, null, null, null, "SOIL_FTG_HDR", "SOIL_EXC"));
            foreach (var f in buriedFtg)
            {
                int r = NextRow(rows);
                int d = f.DataRow;
                string lThk = ResolveLeanThkExpr(f);
                rows.Add(MakeRow("DATA", 2, $"{f.ElementId}::EXC_FTG", "SOIL_EXC",
                    f.Mark, f.Description,
                    $"=QS_Data!G{d}*1000&\"x\"&QS_Data!H{d}*1000&\"x\"&QS_Data!I{d}*1000", "m3",
                    $"=QS_Data!F{d}",
                    $"=QS_Data!G{d}+2*SPEC_Exc_WorkSpace_Ftg",
                    $"=QS_Data!H{d}+2*SPEC_Exc_WorkSpace_Ftg",
                    $"=MAX(0,SPEC_Elev_Natural-SPEC_Topsoil_Thk-(QS_Data!K{d}-({lThk})))",
                    "=SPEC_Exc_SlopeK", null, null,
                    $"=PRODUCT(E{r}:K{r})", $"{f.ElementId}::EXC_FTG", "SOIL_EXC"));
            }
        }

        // Ground beam trench excavation
        if (buriedGb.Count > 0)
        {
            rows.Add(MakeRow("SUB_HDR", 1, "SOIL_GB_HDR", "SOIL_EXC", "", "Ground beam trench excavation / Đào rãnh giằng móng", "", "m3", null, null, null, null, null, null, null, null, "SOIL_GB_HDR", "SOIL_EXC"));
            foreach (var gb in buriedGb)
            {
                int r = NextRow(rows);
                int d = gb.DataRow;
                string lThk = ResolveLeanThkExpr(gb);
                rows.Add(MakeRow("DATA", 2, $"{gb.ElementId}::EXC_GB", "SOIL_EXC",
                    gb.Mark, gb.Description,
                    $"=QS_Data!H{d}*1000&\"x\"&QS_Data!I{d}*1000", "m3",
                    $"=QS_Data!F{d}",
                    $"=QS_Data!G{d}+2*SPEC_Exc_WorkSpace_Beam",
                    $"=QS_Data!H{d}+2*SPEC_Exc_WorkSpace_Beam",
                    $"=MAX(0,SPEC_Elev_Natural-SPEC_Topsoil_Thk-(QS_Data!K{d}-({lThk})))",
                    "=SPEC_Exc_SlopeK", null, null,
                    $"=PRODUCT(E{r}:K{r})", $"{gb.ElementId}::EXC_GB", "SOIL_EXC"));
            }
        }
        AppendManualRows(rows, manualMap, "SOIL_EXC");
        int excSubEnd = NextRow(rows) - 1;
        rows[excTotalRow - 1].Values[11] = RangeSubtotal(excSubStart, excSubEnd);

        // Machine vs Labor Excavation Summary (matches BHFLEX format cleanly!)
        int rMach = NextRow(rows);
        rows.Add(MakeRow("ITEM_HDR", 1, "SOIL_EXC_MACH", "SOIL_SUM", "1.1", "Soil excavation, used machine / Đào đất bằng máy (90%)", "", "m3",
            null, null, null, null, null, null, null, $"=L{excTotalRow}*SPEC_Exc_MachinePct", "SOIL_EXC_MACH", "SOIL_SUM"));
        int rLab = NextRow(rows);
        rows.Add(MakeRow("ITEM_HDR", 1, "SOIL_EXC_LAB", "SOIL_SUM", "1.2", "Soil excavation, used labor / Đào đất sửa hố thủ công (10%)", "", "m3",
            null, null, null, null, null, null, null, $"=L{excTotalRow}*SPEC_Exc_LaborPct", "SOIL_EXC_LAB", "SOIL_SUM"));

        // Occupied underground volume (sum of buried elements directly from QS_Data rows)
        string buriedConcreteExpr = BuildDirectVolumeSumExpr(buriedFtg.Concat(buriedGb).ToList(), true);

        // Backfill reusing excavated soil
        int rReuse = NextRow(rows);
        rows.Add(MakeRow("ITEM_HDR", 1, "SOIL_BACKFILL_REUSE", "SOIL_BACKFILL", "2", "Backfill soil & compaction, K95 (reused) / Lấp đất tận dụng K95", "K=0.95", "m3",
            null, null, null, null, null, null, null, $"=MAX(0,L{excTotalRow}-{buriedConcreteExpr})", "SOIL_BACKFILL_REUSE", "SOIL_BACKFILL"));

        // Soil removal
        int rRemove = NextRow(rows);
        rows.Add(MakeRow("ITEM_HDR", 1, "SOIL_REMOVE", "SOIL_BACKFILL", "3", "Soil removal to spoil area / Vận chuyển đất thừa ra ngoài", "", "m3",
            null, null, null, null, null, null, null, $"=MAX(0,L{excTotalRow}-L{rReuse})", "SOIL_REMOVE", "SOIL_BACKFILL"));

        // Compaction K95 (m2), Base stone K98 (m3), PE film (m2) for SlabOnGrade & Footings
        if (sogList.Count > 0 || buriedFtg.Count > 0)
        {
            int hdrComp = NextRow(rows);
            rows.Add(MakeRow("ITEM_HDR", 1, "SOIL_COMPACT_HDR", "SOIL_COMPACT", "4", "Existing soil compaction K95 / Đầm chặt đất hiện trạng K95", "K=0.95", "m2", null, null, null, null, null, null, null, 0, "SOIL_COMPACT_HDR", "SOIL_COMPACT"));
            int cs = NextRow(rows);
            foreach (var s in sogList.Concat(buriedFtg))
            {
                int r = NextRow(rows);
                int d = s.DataRow;
                rows.Add(MakeRow("DATA", 2, $"{s.ElementId}::COMP95", "SOIL_COMPACT",
                    s.Mark, "Compaction K95 - " + s.Mark, "", "m2",
                    $"=QS_Data!F{d}", $"=QS_Data!G{d}", $"=QS_Data!H{d}", null, null, null, null, $"=PRODUCT(E{r}:K{r})", $"{s.ElementId}::COMP95", "SOIL_COMPACT"));
            }
            AppendManualRows(rows, manualMap, "SOIL_COMPACT");
            int ce = NextRow(rows) - 1;
            rows[hdrComp - 1].Values[11] = RangeSubtotal(cs, ce);
        }

        if (sogList.Count > 0)
        {
            int hdrAgg = NextRow(rows);
            rows.Add(MakeRow("ITEM_HDR", 1, "SOIL_AGG_HDR", "SOIL_AGG", "5", "Aggregate base stone well compacted K98 / Đá cấp phối đầm chặt K98", "Hard cubic", "m3", null, null, null, null, null, null, null, 0, "SOIL_AGG_HDR", "SOIL_AGG"));
            int ags = NextRow(rows);
            foreach (var s in sogList)
            {
                int r = NextRow(rows);
                int d = s.DataRow;
                rows.Add(MakeRow("DATA", 2, $"{s.ElementId}::AGG98", "SOIL_AGG",
                    s.Mark, "AGG Base K98 (250thk) - " + s.Mark, "", "m3",
                    $"=QS_Data!F{d}", $"=QS_Data!G{d}", $"=QS_Data!H{d}", 0.25, "=SPEC_Exc_SlopeK", null, null, $"=PRODUCT(E{r}:K{r})", $"{s.ElementId}::AGG98", "SOIL_AGG"));
            }
            int age = NextRow(rows) - 1;
            rows[hdrAgg - 1].Values[11] = RangeSubtotal(ags, age);

            int hdrPe = NextRow(rows);
            rows.Add(MakeRow("ITEM_HDR", 1, "SOIL_PE_HDR", "SOIL_PE", "6", "PE film 2 plies T0.2mm / Màng PE lót nền 2 lớp", "T0.2mm", "m2", null, null, null, null, null, null, null, 0, "SOIL_PE_HDR", "SOIL_PE"));
            int pes = NextRow(rows);
            foreach (var s in sogList)
            {
                int r = NextRow(rows);
                int d = s.DataRow;
                rows.Add(MakeRow("DATA", 2, $"{s.ElementId}::PE", "SOIL_PE",
                    s.Mark, "PE Film - " + s.Mark, "", "m2",
                    $"=QS_Data!F{d}", $"=QS_Data!G{d}", $"=QS_Data!H{d}", null, "=SPEC_PE_Overlap", null, null, $"=PRODUCT(E{r}:K{r})", $"{s.ElementId}::PE", "SOIL_PE"));
            }
            int pee = NextRow(rows) - 1;
            rows[hdrPe - 1].Values[11] = RangeSubtotal(pes, pee);
        }

        rows.Add(MakeRow("BLANK", 0, "", "SOIL"));
    }

    private void BuildFormworkSection(List<SheetCellRowDto> rows, List<ElementDataDto> elements, Dictionary<string, List<ExistingDetailRowDto>> manualMap)
    {
        rows.Add(MakeRow("WORK_HDR", 0, "SEC_FW", "FW", "III", "FORMWORK WORKS / CÔNG TÁC VÁN KHUÔN (CỐP PHA)", "", "", null, null, null, null, null, null, null, null, "SEC_FW", "FW"));

        // 1. Formwork for Lean Concrete
        var leanEls = elements.Where(e => IsGroup(e, "Foundation", "GroundBeam", "Tank", "Manhole")).ToList();
        if (leanEls.Count > 0 || manualMap.ContainsKey("FW_LEAN"))
        {
            int hdr = NextRow(rows);
            rows.Add(MakeRow("ITEM_HDR", 1, "FW_LEAN_HDR", "FW_LEAN", "1", "Formwork for lean concrete / Ván khuôn bê tông lót", "", "m2", null, null, null, null, null, null, null, 0, "FW_LEAN_HDR", "FW_LEAN"));
            int s = NextRow(rows);
            foreach (var el in leanEls)
            {
                int r = NextRow(rows);
                int d = el.DataRow;
                string lThk = ResolveLeanThkExpr(el);
                string lExt = ResolveLeanExtExpr(el);
                rows.Add(MakeRow("DATA", 2, $"{el.ElementId}::FW_LEAN", "FW_LEAN",
                    el.Mark, "FW Lean " + el.Mark,
                    $"=ROUND(F{r}*1000,0)&\"x\"&ROUND(G{r}*1000,0)&\"x\"&ROUND(H{r}*1000,0)", "m2",
                    $"=QS_Data!F{d}",
                    $"=QS_Data!G{d}+2*({lExt})",
                    $"=QS_Data!H{d}+2*({lExt})",
                    $"={lThk}", null, null, null,
                    $"=E{r}*2*(F{r}+G{r})*H{r}", $"{el.ElementId}::FW_LEAN", "FW_LEAN"));
            }
            AppendManualRows(rows, manualMap, "FW_LEAN");
            int e = NextRow(rows) - 1;
            rows[hdr - 1].Values[11] = RangeSubtotal(s, e);
        }

        // 2. Substructure Formwork
        var subEls = elements.Where(e => IsBucket(e, "Substructure") && !IsGroup(e, "Pile")).ToList();
        if (subEls.Count > 0 || manualMap.ContainsKey("FW_SUB"))
        {
            int hdr = NextRow(rows);
            rows.Add(MakeRow("ITEM_HDR", 1, "FW_SUB_HDR", "FW_SUB", "2", "Formwork for foundation, ground beam, ground slab / Ván khuôn móng, giằng, nền", "Plywood 15mm", "m2", null, null, null, null, null, null, null, 0, "FW_SUB_HDR", "FW_SUB"));
            int s = NextRow(rows);
            foreach (var el in subEls)
            {
                AddFormworkRow(rows, el, "FW_SUB");
            }
            AppendManualRows(rows, manualMap, "FW_SUB");
            int e = NextRow(rows) - 1;
            rows[hdr - 1].Values[11] = RangeSubtotal(s, e);
        }

        // 3. Superstructure Formwork
        var supEls = elements.Where(e => IsBucket(e, "Superstructure")).ToList();
        if (supEls.Count > 0 || manualMap.ContainsKey("FW_SUP"))
        {
            int hdr = NextRow(rows);
            rows.Add(MakeRow("ITEM_HDR", 1, "FW_SUP_HDR", "FW_SUP", "3", "Formwork for column, wall, beam, slab / Ván khuôn cột, vách, dầm, sàn thân", "Plywood 18mm", "m2", null, null, null, null, null, null, null, 0, "FW_SUP_HDR", "FW_SUP"));
            int s = NextRow(rows);
            foreach (var el in supEls)
            {
                AddFormworkRow(rows, el, "FW_SUP");
            }
            AppendManualRows(rows, manualMap, "FW_SUP");
            int e = NextRow(rows) - 1;
            rows[hdr - 1].Values[11] = RangeSubtotal(s, e);
        }

        // 4. Misc / Stair / Lintel / Civil Formwork
        var miscEls = elements.Where(e => !IsGroup(e, "Pile") && !IsBucket(e, "Substructure") && !IsBucket(e, "Superstructure")).ToList();
        if (miscEls.Count > 0 || manualMap.ContainsKey("FW_MISC"))
        {
            int hdr = NextRow(rows);
            rows.Add(MakeRow("ITEM_HDR", 1, "FW_MISC_HDR", "FW_MISC", "4", "Formwork for staircase, lintel, civil structures / VK cầu thang, lanh tô, hố ga", "", "m2", null, null, null, null, null, null, null, 0, "FW_MISC_HDR", "FW_MISC"));
            int s = NextRow(rows);
            foreach (var el in miscEls)
            {
                AddFormworkRow(rows, el, "FW_MISC");
            }
            AppendManualRows(rows, manualMap, "FW_MISC");
            int e = NextRow(rows) - 1;
            rows[hdr - 1].Values[11] = RangeSubtotal(s, e);
        }

        rows.Add(MakeRow("BLANK", 0, "", "FW"));
    }

    private void AddFormworkRow(List<SheetCellRowDto> rows, ElementDataDto el, string secKey)
    {
        int r = NextRow(rows);
        int d = el.DataRow;
        string qtyFormula;
        if (IsGroup(el, "Beam"))
        {
            // Bottom (W*L) + 2 sides (2*H*L) minus clash
            qtyFormula = $"=MAX(0,E{r}*(G{r}+2*H{r})*F{r}-QS_Data!N{d})";
        }
        else if (IsGroup(el, "Slab"))
        {
            // Slab bottom (L*W) + edge (2*(L+W)*t) minus clash
            qtyFormula = $"=MAX(0,E{r}*(F{r}*G{r}+2*(F{r}+G{r})*H{r})-QS_Data!N{d})";
        }
        else if (IsGroup(el, "Stair"))
        {
            // Underside slope (SQRT(L^2+H^2)*W) + stringers & risers allowance
            qtyFormula = $"=MAX(0,E{r}*(SQRT(F{r}^2+H{r}^2)*G{r}+F{r}*H{r})-QS_Data!N{d})";
        }
        else
        {
            // Perimeter sides: Foundation, GroundBeam, Column, Wall, SlabOnGrade edge, Lintel, Manhole, Tank
            qtyFormula = $"=MAX(0,E{r}*2*(F{r}+G{r})*H{r}-QS_Data!N{d})";
        }

        rows.Add(MakeRow("DATA", 2, $"{el.ElementId}::{secKey}", secKey,
            el.Mark, el.Description,
            $"=QS_Data!G{d}*1000&\"x\"&QS_Data!H{d}*1000&\"x\"&QS_Data!I{d}*1000", "m2",
            $"=QS_Data!F{d}", $"=QS_Data!G{d}", $"=QS_Data!H{d}", $"=QS_Data!I{d}",
            null, null, null, qtyFormula, $"{el.ElementId}::{secKey}", secKey));
    }

    private void BuildConcreteSection(List<SheetCellRowDto> rows, List<ElementDataDto> elements, Dictionary<string, List<ExistingDetailRowDto>> manualMap)
    {
        rows.Add(MakeRow("WORK_HDR", 0, "SEC_CONC", "CONC", "IV", "CONCRETE WORKS / CÔNG TÁC BÊ TÔNG", "", "", null, null, null, null, null, null, null, null, "SEC_CONC", "CONC"));

        // 1. Lean Concrete
        var leanEls = elements.Where(e => IsGroup(e, "Foundation", "GroundBeam", "SlabOnGrade", "Tank", "Manhole")).ToList();
        if (leanEls.Count > 0 || manualMap.ContainsKey("CONC_LEAN"))
        {
            int hdr = NextRow(rows);
            rows.Add(MakeRow("ITEM_HDR", 1, "CONC_LEAN_HDR", "CONC_LEAN", "1", "Lean concrete for foundation, ground beam & slab / Bê tông lót M100#", "M100#", "m3", null, null, null, null, null, null, null, 0, "CONC_LEAN_HDR", "CONC_LEAN"));
            int s = NextRow(rows);
            foreach (var el in leanEls)
            {
                int r = NextRow(rows);
                int d = el.DataRow;
                string lThk = ResolveLeanThkExpr(el);
                string lExt = ResolveLeanExtExpr(el);
                string lF = $"=QS_Data!G{d}+2*({lExt})";
                string wF = $"=QS_Data!H{d}+2*({lExt})";
                string hF = $"={lThk}";
                rows.Add(MakeRow("DATA", 2, $"{el.ElementId}::CONC_LEAN", "CONC_LEAN",
                    el.Mark, "Lean " + el.Mark,
                    $"=ROUND(F{r}*1000,0)&\"x\"&ROUND(G{r}*1000,0)&\"x\"&ROUND(H{r}*1000,0)", "m3",
                    $"=QS_Data!F{d}", lF, wF, hF, "=SPEC_Conc_Waste", null, null,
                    $"=PRODUCT(E{r}:K{r})", $"{el.ElementId}::CONC_LEAN", "CONC_LEAN"));
            }
            AppendManualRows(rows, manualMap, "CONC_LEAN");
            int e = NextRow(rows) - 1;
            rows[hdr - 1].Values[11] = RangeSubtotal(s, e);
        }

        // 2. Substructure Concrete
        var subEls = elements.Where(e => IsBucket(e, "Substructure") && !IsGroup(e, "Pile")).ToList();
        if (subEls.Count > 0 || manualMap.ContainsKey("CONC_SUB"))
        {
            int hdr = NextRow(rows);
            rows.Add(MakeRow("ITEM_HDR", 1, "CONC_SUB_HDR", "CONC_SUB", "2", "Concrete for foundation, ground beam, ground slab / BT móng, giằng, nền", "Grade M400#", "m3", null, null, null, null, null, null, null, 0, "CONC_SUB_HDR", "CONC_SUB"));
            int s = NextRow(rows);
            foreach (var el in subEls) AddConcreteRow(rows, el, "CONC_SUB");
            AppendManualRows(rows, manualMap, "CONC_SUB");
            int e = NextRow(rows) - 1;
            rows[hdr - 1].Values[11] = RangeSubtotal(s, e);
        }

        // 3. Superstructure Concrete
        var supEls = elements.Where(e => IsBucket(e, "Superstructure")).ToList();
        if (supEls.Count > 0 || manualMap.ContainsKey("CONC_SUP"))
        {
            int hdr = NextRow(rows);
            rows.Add(MakeRow("ITEM_HDR", 1, "CONC_SUP_HDR", "CONC_SUP", "3", "Concrete for column, wall, beam, slab / BT cột, vách, dầm, sàn thân", "Grade M400#", "m3", null, null, null, null, null, null, null, 0, "CONC_SUP_HDR", "CONC_SUP"));
            int s = NextRow(rows);
            foreach (var el in supEls) AddConcreteRow(rows, el, "CONC_SUP");
            AppendManualRows(rows, manualMap, "CONC_SUP");
            int e = NextRow(rows) - 1;
            rows[hdr - 1].Values[11] = RangeSubtotal(s, e);
        }

        // 4. Misc / Stair / Lintel / Civil Concrete
        var miscEls = elements.Where(e => !IsGroup(e, "Pile") && !IsBucket(e, "Substructure") && !IsBucket(e, "Superstructure")).ToList();
        if (miscEls.Count > 0 || manualMap.ContainsKey("CONC_MISC"))
        {
            int hdr = NextRow(rows);
            rows.Add(MakeRow("ITEM_HDR", 1, "CONC_MISC_HDR", "CONC_MISC", "4", "Concrete for staircase, lintel, civil structures / BT cầu thang, lanh tô, hố ga", "Grade M350#", "m3", null, null, null, null, null, null, null, 0, "CONC_MISC_HDR", "CONC_MISC"));
            int s = NextRow(rows);
            foreach (var el in miscEls) AddConcreteRow(rows, el, "CONC_MISC");
            AppendManualRows(rows, manualMap, "CONC_MISC");
            int e = NextRow(rows) - 1;
            rows[hdr - 1].Values[11] = RangeSubtotal(s, e);
        }

        // 5. Industrial Slab Polishing & Joint Cutting
        var slabs = elements.Where(e => IsGroup(e, "SlabOnGrade", "Slab")).ToList();
        if (slabs.Count > 0)
        {
            int hdrPol = NextRow(rows);
            rows.Add(MakeRow("ITEM_HDR", 1, "CONC_POLISH_HDR", "CONC_POLISH", "5", "Polishing & power float concrete slab / Xoa nền đánh bóng mặt bê tông sàn", "", "m2", null, null, null, null, null, null, null, 0, "CONC_POLISH_HDR", "CONC_POLISH"));
            int ps = NextRow(rows);
            foreach (var sl in slabs)
            {
                int r = NextRow(rows);
                int d = sl.DataRow;
                rows.Add(MakeRow("DATA", 2, $"{sl.ElementId}::POLISH", "CONC_POLISH",
                    sl.Mark, "Power float " + sl.Mark, "", "m2",
                    $"=QS_Data!F{d}", $"=QS_Data!G{d}", $"=QS_Data!H{d}", null, null, null, null,
                    $"=PRODUCT(E{r}:K{r})", $"{sl.ElementId}::POLISH", "CONC_POLISH"));
            }
            int pe = NextRow(rows) - 1;
            rows[hdrPol - 1].Values[11] = RangeSubtotal(ps, pe);

            int hdrJnt = NextRow(rows);
            rows.Add(MakeRow("ITEM_HDR", 1, "CONC_CUTJNT_HDR", "CONC_CUTJNT", "6", "Cutting contraction joint on slab (grid) / Cắt mạch khe co giãn sàn", "h=15mm", "m", null, null, null, null, null, null, null, 0, "CONC_CUTJNT_HDR", "CONC_CUTJNT"));
            int js = NextRow(rows);
            foreach (var sl in slabs.Where(x => IsGroup(x, "SlabOnGrade")))
            {
                int r = NextRow(rows);
                int d = sl.DataRow;
                rows.Add(MakeRow("DATA", 2, $"{sl.ElementId}::CUTJNT", "CONC_CUTJNT",
                    sl.Mark, "Cut joint grid - " + sl.Mark, "", "m",
                    $"=QS_Data!F{d}", $"=QS_Data!G{d}", $"=QS_Data!H{d}", null, null, null, null,
                    $"=E{r}*(INT(F{r}/SPEC_Joint_Grid)*G{r}+INT(G{r}/SPEC_Joint_Grid)*F{r})", $"{sl.ElementId}::CUTJNT", "CONC_CUTJNT"));
            }
            int je = NextRow(rows) - 1;
            if (je >= js) rows[hdrJnt - 1].Values[11] = RangeSubtotal(js, je);
        }
    }

    private void AddConcreteRow(List<SheetCellRowDto> rows, ElementDataDto el, string secKey)
    {
        int r = NextRow(rows);
        int d = el.DataRow;
        string qtyFormula;
        if (IsGroup(el, "Stair"))
        {
            // Slope slab + triangular steps approx cleanly as Excel formula
            qtyFormula = $"=MAX(0,E{r}*((SQRT(F{r}^2+H{r}^2)*G{r}*0.15)+(0.5*F{r}*H{r}*G{r}*0.25))*I{r}-QS_Data!M{d})";
        }
        else
        {
            qtyFormula = $"=MAX(0,PRODUCT(E{r}:K{r})-QS_Data!M{d})";
        }
        rows.Add(MakeRow("DATA", 2, $"{el.ElementId}::{secKey}", secKey,
            el.Mark, el.Description,
            $"=QS_Data!G{d}*1000&\"x\"&QS_Data!H{d}*1000&\"x\"&QS_Data!I{d}*1000&\" (\"&QS_Data!L{d}&\")\"", "m3",
            $"=QS_Data!F{d}", $"=QS_Data!G{d}", $"=QS_Data!H{d}", $"=QS_Data!I{d}",
            "=SPEC_Conc_Waste", null, null, qtyFormula, $"{el.ElementId}::{secKey}", secKey));
    }

    private static string BuildDirectVolumeSumExpr(List<ElementDataDto> els, bool includeLean)
    {
        if (els.Count == 0) return "0";
        var parts = els.Select(e =>
        {
            int d = e.DataRow;
            string lThk = ResolveLeanThkExpr(e);
            string lExt = ResolveLeanExtExpr(e);
            return includeLean
                ? $"(QS_Data!F{d}*(QS_Data!G{d}*QS_Data!H{d}*QS_Data!I{d}+(QS_Data!G{d}+2*({lExt}))*(QS_Data!H{d}+2*({lExt}))*({lThk})))"
                : $"(QS_Data!F{d}*QS_Data!G{d}*QS_Data!H{d}*QS_Data!I{d})";
        });
        return "(" + string.Join("+", parts) + ")";
    }

    private static string ResolveLeanThkExpr(ElementDataDto el)
    {
        int d = el.DataRow;
        string defName = (el.Group ?? "").Trim().ToLowerInvariant() switch
        {
            "groundbeam" or "beam" => "SPEC_Lean_Beam_Thk",
            "slabongrade" or "slab" => "SPEC_Lean_Slab_Thk",
            "tank" or "manhole" => "SPEC_Lean_Tank_Thk",
            _ => "SPEC_Lean_Ftg_Thk"
        };
        return $"IF(QS_Data!P{d}=\"\",{defName},QS_Data!P{d})";
    }

    private static string ResolveLeanExtExpr(ElementDataDto el)
    {
        int d = el.DataRow;
        string defName = (el.Group ?? "").Trim().ToLowerInvariant() switch
        {
            "groundbeam" or "beam" => "SPEC_Lean_Beam_Ext",
            "slabongrade" or "slab" => "SPEC_Lean_Slab_Ext",
            "tank" or "manhole" => "SPEC_Lean_Tank_Ext",
            _ => "SPEC_Lean_Ftg_Ext"
        };
        return $"IF(QS_Data!Q{d}=\"\",{defName},QS_Data!Q{d})";
    }

    private static Dictionary<string, List<ExistingDetailRowDto>> GroupManualRows(IReadOnlyList<ExistingDetailRowDto>? existing)
    {
        var dict = new Dictionary<string, List<ExistingDetailRowDto>>(StringComparer.OrdinalIgnoreCase);
        if (existing == null) return dict;
        foreach (var row in existing)
        {
            bool isManual = string.Equals(row.TagId, "MANUAL", StringComparison.OrdinalIgnoreCase)
                         || (string.IsNullOrWhiteSpace(row.TagId) && !string.IsNullOrWhiteSpace(row.Description));
            if (!isManual) continue;
            string sec = string.IsNullOrWhiteSpace(row.SectionKey) ? "CONC_MISC" : row.SectionKey.Trim();
            if (!dict.TryGetValue(sec, out var list))
            {
                list = new List<ExistingDetailRowDto>();
                dict[sec] = list;
            }
            list.Add(row);
        }
        return dict;
    }

    private static void AppendManualRows(List<SheetCellRowDto> rows, Dictionary<string, List<ExistingDetailRowDto>> manualMap, string sectionKey)
    {
        if (!manualMap.TryGetValue(sectionKey, out var list)) return;
        foreach (var m in list)
        {
            rows.Add(MakeRow("MANUAL", 2, "MANUAL", sectionKey,
                m.Code, m.Description, m.SpecText, m.Unit,
                NullIfEmpty(m.NoExpr), NullIfEmpty(m.LExpr), NullIfEmpty(m.WExpr), NullIfEmpty(m.HExpr),
                NullIfEmpty(m.AExpr), NullIfEmpty(m.BExpr), NullIfEmpty(m.CExpr),
                string.IsNullOrWhiteSpace(m.QtyExpr) ? $"=PRODUCT(E{NextRow(rows)}:K{NextRow(rows)})" : m.QtyExpr,
                "MANUAL", sectionKey));
        }
    }

    private static object? NullIfEmpty(string s) => string.IsNullOrWhiteSpace(s) ? null : s;

    private static string RangeSubtotal(int startRow, int endRow) => endRow >= startRow ? $"=SUBTOTAL(9,L{startRow}:L{endRow})" : "0";

    private static int NextRow(List<SheetCellRowDto> rows) => rows.Count + 1;

    private static bool IsGroup(ElementDataDto el, params string[] names) =>
        names.Any(n => string.Equals(el.Group?.Trim(), n, StringComparison.OrdinalIgnoreCase));

    private static bool IsBucket(ElementDataDto el, string bucket) =>
        string.Equals(el.BoqBucket?.Trim(), bucket, StringComparison.OrdinalIgnoreCase);

    private static bool IsBuried(ElementDataDto el) =>
        string.Equals(el.IsBuried?.Trim(), "Y", StringComparison.OrdinalIgnoreCase) || IsGroup(el, "Foundation", "GroundBeam");

    private static string Fmt(double v) => v.ToString("0.###", CultureInfo.InvariantCulture);

    private static string EscapeExcel(string s) => (s ?? string.Empty).Replace("\"", "\"\"");

    public static string SanitizeSheetName(string name)
    {
        var invalid = new[] { '\\', '/', '?', '*', '[', ']', ':' };
        string cleaned = new string((name ?? "Sheet").Select(ch => invalid.Contains(ch) ? '_' : ch).ToArray());
        return cleaned.Length > 31 ? cleaned.Substring(0, 31) : cleaned;
    }

    public static SheetCellRowDto MakeRow(string kind, int outlineLevel, string tagId, string sectionKey, params object?[] cols)
    {
        var list = new List<object?>(TotalCols);
        for (int i = 0; i < TotalCols; i++)
        {
            list.Add(i < cols.Length ? cols[i] : null);
        }
        return new SheetCellRowDto
        {
            RowKind = kind,
            OutlineLevel = outlineLevel,
            TagId = tagId,
            SectionKey = sectionKey,
            Values = list
        };
    }
}
