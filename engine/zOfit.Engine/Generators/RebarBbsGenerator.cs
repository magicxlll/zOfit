using System.Globalization;
using ZOfit.Engine.Models;

namespace ZOfit.Engine.Generators;

/// <summary>
/// Generates standard TKT_* Rebar Bending Schedule rows driven by Master (QS_Data) + Detail (QS_RebarConfig v3.0).
/// Implements transparent parametric formulas (BarDir + DistDir + HookL/Bv/HookR + SL.m/L.m overrides + +/-Pb/+/-Lth)
/// compatible with Excel 2007-365 (no LET/IFS/XLOOKUP) and 1-to-1 mapped to 41 BarShapes (D01..D41).
/// </summary>
public sealed class RebarBbsGenerator
{
    public const int FirstDataRow = 11;
    public const int TotalCols = 26; // A..X (1..24) + Y(TagId, 25) + Z(BoqBucket, 26)

    public ExpandRebarPresetsResponse ExpandPresetsToConfigs(ExpandRebarPresetsRequest req)
    {
        var configuredIds = new HashSet<string>(
            (req.ExistingConfigs ?? new List<RebarConfigRowDto>())
                .Where(c => !string.Equals(c.RowKind, "BLK_HDR", StringComparison.OrdinalIgnoreCase)
                         && !string.Equals(c.RowKind, "BLK_END", StringComparison.OrdinalIgnoreCase))
                .Select(c => c.ElementId?.Trim() ?? string.Empty)
                .Where(s => !string.IsNullOrEmpty(s)),
            StringComparer.OrdinalIgnoreCase);

        var added = new List<RebarConfigRowDto>();
        foreach (var el in req.Elements)
        {
            if (string.IsNullOrWhiteSpace(el.ElementId) || configuredIds.Contains(el.ElementId.Trim()))
                continue;
            if (string.IsNullOrWhiteSpace(el.RebarTemplate) || string.Equals(el.RebarTemplate.Trim(), "NONE", StringComparison.OrdinalIgnoreCase))
                continue;

            added.AddRange(ExpandSingleElementToConfigRows(el));
        }

        return new ExpandRebarPresetsResponse
        {
            AddedRowsCount = added.Count,
            Rows = added
        };
    }

    public List<RebarConfigRowDto> ExpandSingleElementToConfigRows(ElementDataDto el)
    {
        var rows = new List<RebarConfigRowDto>();
        string tpl = (el.RebarTemplate ?? "").Trim().ToUpperInvariant();
        string mk = (el.Mark ?? el.ElementId).Trim();
        string coverExpr = "=" + ResolveCoverName(el.Group);

        void AddBar(
            string barSuffix, string role, string shape, double dia, double layerMult,
            string barDir, string distDir, string hookL, string hookR,
            double? sp1, double? sp2, string ovQty, string ovLen, string dPb, string dLth,
            string note, double? le1a = null, double? ri2a = null)
        {
            if (dia <= 0) return;
            rows.Add(new RebarConfigRowDto
            {
                ConfigRow = 0,
                RowKind = "BAR",
                ElementId = el.ElementId,
                Building = el.Building,
                Group = el.Group,
                BarId = $"{mk}-{barSuffix}",
                Role = role,
                ShapeCode = shape,
                Dia = dia,
                LayerMult = layerMult > 0 ? layerMult : 1,
                BarDir = barDir,
                DistDir = distDir,
                HookLExpr = hookL,
                CoverExpr = coverExpr,
                HookRExpr = hookR,
                Space1Mm = sp1,
                Space2Mm = sp2,
                OverrideQtyExpr = ovQty,
                OverrideLenExpr = ovLen,
                DeltaDistExpr = dPb,
                DeltaLenExpr = dLth,
                Le1aMm = le1a,
                Ri2aMm = ri2a,
                CountOrSpace = !string.IsNullOrWhiteSpace(ovQty) ? ovQty : (sp2 is > 0 ? $"{Fmt(sp1 ?? 150)}/{Fmt(sp2.Value)}" : $"@{Fmt(sp1 ?? 150)}"),
                DistMode = distDir,
                LenMode = barDir,
                Note = note
            });
        }

        double dMainX = el.MainXDia > 0 ? el.MainXDia : 16;
        double dMainY = el.MainYDia > 0 ? el.MainYDia : 14;
        double dTopX = el.TopXDia > 0 ? el.TopXDia : dMainX;
        double dTopY = el.TopYDia > 0 ? el.TopYDia : dMainY;
        double dStir = el.StirrupDia > 0 ? el.StirrupDia : 10;

        double mainQtyOrSpaceX = ParseNumOr(el.MainXSpec, 4);
        double spMainX = ParseNumOr(el.MainXSpec, 150);
        double spMainY = ParseNumOr(el.MainYSpec, 150);
        var (stir1, stir2) = ParseDualSpacing(el.StirrupSpec, 150, 200);

        switch (tpl)
        {
            case "BEAM_FULL":
            case "BEAM_STD":
                AddBar("T1", "BEAM_TOP", "D07", dTopX, 1, "X", "FIXED", "20", "20", null, null,
                    IsFixedSpec(el.TopXSpec) ? el.TopXSpec : "4", "", "", "", "Thep chu mien tren (mo 20D)");
                AddBar("B1", "BEAM_BOT", "D07", dMainX, 1, "X", "FIXED", "20", "20", null, null,
                    IsFixedSpec(el.MainXSpec) ? el.MainXSpec : Fmt(mainQtyOrSpaceX), "", "", "", "Thep chu mien duoi (mo 20D)");
                if (tpl == "BEAM_FULL" || el.HeightM >= 0.70)
                    AddBar("SK", "BEAM_SKIN", "D01", 12, 2, "X", "FIXED", "0", "0", null, null,
                        "1", "", "", "=2*MAX(250,SPEC_Lan_Tension*L{r})", "Thep gia thanh dam 2 mat (1x2=2 cay)");
                if (tpl == "BEAM_FULL")
                {
                    AddBar("GG", "BEAM_SUP_ADD", "D06", dTopX, 2, "X*0.25", "FIXED", "20", "0", null, null,
                        "2", "", "", "=SPEC_Lan_Tension*L{r}", "Gia cuong goi (L/4 x 2 dau)");
                    AddBar("GN", "BEAM_MID_ADD", "D01", dMainX, 1, "X*0.7", "FIXED", "0", "0", null, null,
                        "2", "", "", "", "Gia cuong bung nhip (0.7L)");
                }
                AddBar("ST1", "STIRRUP_RECT", "D39", dStir, 1, "PERIM_YH", "ZONE2_X", "10", "10", stir1, stir2,
                    "", "", "", "", $"Dai kin chu nhat (@{Fmt(stir1)} goi / @{Fmt(stir2)} giua)");
                if (tpl == "BEAM_FULL" || el.WidthM >= 0.40)
                    AddBar("ST2", "STIRRUP_C", "D24", 8, 1, "TIE_Y", "ZONE2_X", "10", "10", stir1, stir2,
                        "", "", "", "", "Dai C moc giua dam");
                break;

            case "COL_STAGGER":
                AddBar("C1A", "COL_MAIN_A", "D06", dMainX, 1, "H", "FIXED", "30", "0", null, null,
                    "4", "", "", "=SPEC_Llap_Mult*L{r}+SPEC_Col_StaggerOffset", "Thep chu nhom A (noi sole +Stagger)");
                AddBar("C1B", "COL_MAIN_B", "D06", dMainX, 1, "H", "FIXED", "30", "0", null, null,
                    "4", "", "", "=SPEC_Llap_Mult*L{r}", "Thep chu nhom B (noi chuan +45D)");
                AddBar("D1", "STIRRUP_RECT", "D39", dStir, 1, "PERIM_XY", "ZONE2_H", "10", "10", 100, 200,
                    "", "", "", "", "Dai bao ngoai cot (@100 noi / @200 giua)");
                AddBar("D2", "STIRRUP_INNER", "D39", dStir, 1, "PERIM_XY_INNER", "ZONE2_H", "10", "10", 100, 200,
                    "", "", "", "", "Dai long trong cot");
                AddBar("D3", "STIRRUP_C", "D24", 8, 1, "TIE_Y", "ZONE2_H", "10", "10", 100, 200,
                    "", "", "", "", "Dai C cot");
                break;

            case "COL_STD":
                AddBar("C1", "COL_MAIN_B", "D06", dMainX, 1, "H", "FIXED", "30", "0", null, null,
                    IsFixedSpec(el.MainXSpec) ? el.MainXSpec : "4", "", "", "=SPEC_Llap_Mult*L{r}", "Thep chu cot");
                AddBar("D1", "STIRRUP_RECT", "D39", dStir, 1, "PERIM_XY", "ZONE2_H", "10", "10", 100, 200,
                    "", "", "", "", "Dai kin cot");
                break;

            case "WALL_2FACE":
                AddBar("VD", "WALL_VERT", "D06", el.MainXDia > 0 ? el.MainXDia : 14, 2, "H", "X", "40", "0", spMainX, null,
                    "", "", "", "", "Thep dung vach 2 mat (rai doc dai L)");
                AddBar("HN", "WALL_HORZ", "D07", el.MainYDia > 0 ? el.MainYDia : 12, 2, "X", "H", "100", "100", spMainY, null,
                    "", "", "", "", "Thep ngang vach 2 mat (rai doc cao H)");
                AddBar("DC", "WALL_C_TIE", "D24", 8, 1, "TIE_Y", "XH_GRID", "10", "10", 600, 600,
                    "", "", "", "", "Dai C moc giu 2 lop vach (@600x@600)");
                break;

            case "SLAB_1WAY":
                AddBar("S1", "SLAB_BOT_X", "D01", dMainX > 0 ? dMainX : 10, 1, "X", "Y", "0", "0", ParseNumOr(el.MainXSpec, 150), null,
                    "", "", "", "", "San lop duoi phuong X");
                AddBar("S3", "SLAB_TOP_X", "D07", dTopX > 0 ? dTopX : 10, 1, "X", "Y", "100", "100", ParseNumOr(el.TopXSpec, 150), null,
                    "", "", "", "", "San lop tren phuong X");
                break;

            case "SLAB_2WAY_1LAYER":
                AddBar("S1", "SLAB_BOT_X", "D01", dMainX > 0 ? dMainX : 12, 1, "X", "Y", "0", "0", ParseNumOr(el.MainXSpec, 200), null,
                    "", "", "", "", "San lop duoi phuong X");
                AddBar("S2", "SLAB_BOT_Y", "D01", dMainY > 0 ? dMainY : 12, 1, "Y", "X", "0", "0", ParseNumOr(el.MainYSpec, 200), null,
                    "", "", "", "", "San lop duoi phuong Y");
                break;

            case "SLAB_2WAY_2LAYER":
                AddBar("S1", "SLAB_BOT_X", "D01", dMainX > 0 ? dMainX : 12, 1, "X", "Y", "0", "0", ParseNumOr(el.MainXSpec, 200), null,
                    "", "", "", "", "San lop duoi phuong X");
                AddBar("S2", "SLAB_BOT_Y", "D01", dMainY > 0 ? dMainY : 12, 1, "Y", "X", "0", "0", ParseNumOr(el.MainYSpec, 200), null,
                    "", "", "", "", "San lop duoi phuong Y");
                AddBar("S3", "SLAB_TOP_X", "D07", dTopX > 0 ? dTopX : 12, 1, "X", "Y", "100", "100", ParseNumOr(el.TopXSpec, 200), null,
                    "", "", "", "", "San lop tren phuong X");
                AddBar("S4", "SLAB_TOP_Y", "D07", dTopY > 0 ? dTopY : 12, 1, "Y", "X", "100", "100", ParseNumOr(el.TopYSpec, 200), null,
                    "", "", "", "", "San lop tren phuong Y");
                AddBar("CK", "SLAB_CHAIR", "D07", 10, 1, "CHAIR_H", "XY_GRID", "12", "12", 600, 600,
                    "", "", "", "", "Chan cho ke thep lop tren (@600x@600)");
                break;

            case "SLAB_WITH_HATS":
                AddBar("S1", "SLAB_BOT_X", "D01", dMainX > 0 ? dMainX : 12, 1, "X", "Y", "0", "0", 200, null,
                    "", "", "", "", "San lop duoi phuong X");
                AddBar("S2", "SLAB_BOT_Y", "D01", dMainY > 0 ? dMainY : 12, 1, "Y", "X", "0", "0", 200, null,
                    "", "", "", "", "San lop duoi phuong Y");
                AddBar("M1", "SLAB_HAT_X", "D07", dTopX > 0 ? dTopX : 12, 2, "X*0.25", "Y", "100", "100", 150, null,
                    "", "", "", "=SPEC_Lan_Tension*L{r}", "Thep mu goi phuong X (L/4 x 2 canh)");
                AddBar("M2", "SLAB_HAT_Y", "D07", dTopY > 0 ? dTopY : 12, 2, "Y*0.25", "X", "100", "100", 150, null,
                    "", "", "", "=SPEC_Lan_Tension*L{r}", "Thep mu goi phuong Y (W/4 x 2 canh)");
                AddBar("CK", "SLAB_CHAIR", "D07", 10, 1, "CHAIR_H", "XY_GRID", "12", "12", 600, 600,
                    "", "", "", "", "Chan cho ke mu san (@600x@600)");
                break;

            case "FTG_2WAY":
            case "MANHOLE_STD":
                AddBar("M1", "FTG_BOT_X", "D07", el.MainXDia > 0 ? el.MainXDia : 14, 1, "X", "Y", "150", "150", spMainX, null,
                    "", "", "", "", "Thep day phuong X (be mo 150)");
                AddBar("M2", "FTG_BOT_Y", "D07", el.MainYDia > 0 ? el.MainYDia : 14, 1, "Y", "X", "150", "150", spMainY, null,
                    "", "", "", "", "Thep day phuong Y (be mo 150)");
                break;

            case "STAIR_6BAR":
                AddBar("ST-01", "STAIR_ST01", "D15", 12, 1, "SLOPED_XH", "Y", "250", "450", 150, null,
                    "", "", "", "", "Thep chu ban nghieng");
                AddBar("ST-03", "STAIR_ST03", "D06", 10, 1, "X*0.25", "Y", "120", "0", 150, null,
                    "", "", "", "", "Thep mu goi tren ban thang");
                AddBar("ST-04", "STAIR_ST04", "D01", 8, 1, "Y", "X", "0", "0", 200, null,
                    "", "", "", "", "Thep phan bo ngang ban truot");
                break;

            case "PILE_HEAD":
                AddBar("P1", "PILE_DOWEL", "D06", 16, 1, "CUSTOM", "FIXED", "200", "0", null, null,
                    "6", "=SPEC_Pile_CutHead*1000+1000", "", "", "Thep neo dau coc PHC");
                AddBar("P2", "PILE_HOOP", "D41", 8, 1, "CUSTOM", "FIXED", "150", "0", null, null,
                    "1", "=ROUND(2*PI()*150*8,0)", "", "", "Dai xoan tron dau coc");
                break;
        }

        return rows;
    }

    public GeneratedSheetDto Generate(
        string building,
        ProjectSpecDto spec,
        IReadOnlyList<ElementDataDto> elements,
        IReadOnlyList<ExistingTktRowDto>? existingRows = null,
        IReadOnlyList<RebarConfigRowDto>? rebarConfigs = null)
    {
        var rows = new List<SheetCellRowDto>();
        string cleanSheet = RcDetailGenerator.SanitizeSheetName($"TKT_{building}");

        var manualRows = (existingRows ?? Array.Empty<ExistingTktRowDto>())
            .Where(r => string.Equals(r.TagId, "MANUAL", StringComparison.OrdinalIgnoreCase))
            .ToList();

        var cfgByEl = (rebarConfigs ?? Array.Empty<RebarConfigRowDto>())
            .Where(c => !string.IsNullOrWhiteSpace(c.ElementId)
                     && !string.Equals(c.RowKind, "BLK_HDR", StringComparison.OrdinalIgnoreCase)
                     && !string.Equals(c.RowKind, "BLK_END", StringComparison.OrdinalIgnoreCase))
            .GroupBy(c => c.ElementId.Trim(), StringComparer.OrdinalIgnoreCase)
            .ToDictionary(g => g.Key, g => g.ToList(), StringComparer.OrdinalIgnoreCase);

        int stt = 0;
        int currentRow = FirstDataRow;

        foreach (var el in elements.Where(e => string.Equals(e.Building, building, StringComparison.OrdinalIgnoreCase)))
        {
            List<RebarConfigRowDto> barConfigs;
            if (cfgByEl.TryGetValue(el.ElementId, out var explicitList) && explicitList.Count > 0)
            {
                barConfigs = explicitList;
            }
            else if (!string.IsNullOrWhiteSpace(el.RebarTemplate) && !string.Equals(el.RebarTemplate.Trim(), "NONE", StringComparison.OrdinalIgnoreCase))
            {
                barConfigs = ExpandSingleElementToConfigRows(el);
            }
            else
            {
                continue;
            }

            foreach (var cfg in barConfigs)
            {
                stt++;
                rows.Add(BuildTktRowFromConfig(currentRow, stt, el, cfg));
                currentRow++;
            }
        }

        foreach (var m in manualRows)
        {
            stt++;
            rows.Add(BuildManualTktRow(currentRow, stt, m));
            currentRow++;
        }

        return new GeneratedSheetDto
        {
            SheetName = cleanSheet,
            Building = building,
            AddedCount = rows.Count - manualRows.Count,
            UpdatedCount = 0,
            PreservedManualCount = manualRows.Count,
            RemovedFlaggedCount = 0,
            Rows = rows
        };
    }

    private SheetCellRowDto BuildTktRowFromConfig(int r, int stt, ElementDataDto el, RebarConfigRowDto cfg)
    {
        int d = el.DataRow;
        int rc = cfg.ConfigRow; // >0 means backed by QS_RebarConfig v3.0 transparent row!
        string coverName = ResolveCoverName(el.Group);
        var meta = BarShapeCatalog.Get(cfg.ShapeCode);

        // Live links back to QS_Data & QS_RebarConfig v3.0
        object compCell = rc > 0
            ? $"=QS_Data!D{d}&\" - \"&QS_RebarConfig!J{rc}&IF(QS_RebarConfig!Y{rc}<>\"\",\" (\"&QS_RebarConfig!Y{rc}&\")\",\"\")"
            : $"{el.Mark} ({el.Description})";
        object markCell = rc > 0 ? $"=QS_RebarConfig!I{rc}" : cfg.BarId;
        object diaCell = rc > 0 ? $"=QS_RebarConfig!L{rc}" : cfg.Dia;
        object noteCell = rc > 0 ? $"=QS_RebarConfig!Y{rc}" : cfg.Note;

        object?[] segs = rc > 0
            ? MapConfigV3SegmentsToTkt(rc, meta)
            : ResolveFallbackSegments(r, d, coverName, cfg, meta);

        object qtyPerEl = rc > 0
            ? $"=QS_RebarConfig!M{rc}*QS_RebarConfig!AG{rc}"
            : ResolveFallbackQtyPerEl(r, d, coverName, cfg);

        string lenExpr = meta.LenExpr.Replace("{r}", r.ToString(CultureInfo.InvariantCulture));
        string geomLen = meta.BendCount > 0
            ? $"(({lenExpr})-{meta.BendCount}*IFERROR(SPEC_BendDeduct90,IFERROR(CFG_BendDeduct,1))*O{r})"
            : $"({lenExpr})";

        string barLenFormula = meta.HasLap
            ? $"=IF(O{r}=\"\",\"\",ROUND(({geomLen}+IF(AND(O{r}>=10,O{r}<IFERROR(SPEC_CouplerMinD,IFERROR(CFG_CouplerMinD,20))),MAX(0,CEILING({geomLen}/IFERROR(SPEC_StdBarLen,IFERROR(CFG_BarStdLen,11700)),1)-1)*IFERROR(SPEC_Llap_Mult,IFERROR(CFG_LapRatio,45))*O{r},0))/1000,3))"
            : $"=IF(O{r}=\"\",\"\",ROUND({geomLen}/1000,3))";

        object couplerFormula = meta.HasLap
            ? $"=IF(OR(O{r}=\"\",S{r}=\"\"),0,IF(O{r}>=IFERROR(SPEC_CouplerMinD,IFERROR(CFG_CouplerMinD,20)),R{r}*MAX(0,CEILING({geomLen}/IFERROR(SPEC_StdBarLen,IFERROR(CFG_BarStdLen,11700)),1)-1),0))"
            : 0;

        var vals = new object?[TotalCols];
        vals[0] = stt;
        vals[1] = compCell;
        vals[2] = markCell;
        vals[3] = meta.Code;
        vals[4] = meta.Code;
        for (int i = 0; i < 9; i++) vals[5 + i] = segs[i];
        vals[14] = diaCell;
        vals[15] = qtyPerEl;
        vals[16] = $"=QS_Data!F{d}";
        vals[17] = $"=IF(O{r}=\"\",\"\",P{r}*Q{r})";
        vals[18] = barLenFormula;
        vals[19] = couplerFormula;
        vals[20] = $"=IF(O{r}=\"\",\"\",ROUND(R{r}*S{r},2))";
        vals[21] = $"=IF(O{r}=\"\",\"\",ROUND(PI()*(O{r}/1000)^2/4*7850,3))";
        vals[22] = $"=IF(O{r}=\"\",\"\",ROUND(U{r}*V{r},2))";
        vals[23] = noteCell;
        vals[24] = $"{el.ElementId}::{cfg.BarId}";
        vals[25] = el.BoqBucket;

        return new SheetCellRowDto
        {
            RowKind = "DATA",
            OutlineLevel = 0,
            TagId = $"{el.ElementId}::{cfg.BarId}",
            SectionKey = el.BoqBucket,
            Values = vals.ToList()
        };
    }

    /// <summary>
    /// 1-to-1 live link mapping from QS_RebarConfig v3.0 segments (Z=Left, AA=Le.1a, AB=Le.1b, AC=Mid, AD=Ri.2b, AE=Ri.2a, AF=Right)
    /// into TKT columns F..N (L1..L7, M1, M2) matching BarShapeCatalog length expressions.
    /// </summary>
    private static object?[] MapConfigV3SegmentsToTkt(int rc, BarShapeMeta meta)
    {
        var segs = new object?[9]; // 0..6 = F..L (L1..L7), 7 = M (M1), 8 = N (M2)
        string q = "QS_RebarConfig!";
        string sL = $"{q}Z{rc}", sL1a = $"{q}AA{rc}", sL1b = $"{q}AB{rc}", sMid = $"{q}AC{rc}";
        string sR2b = $"{q}AD{rc}", sR2a = $"{q}AE{rc}", sR = $"{q}AF{rc}";

        switch (meta.Code)
        {
            case "D01":
                segs[0] = $"=IF({sL}+{sR}>0,{q}AH{rc},{sMid})";
                break;

            case "D02":
            case "D04":
            case "D24":
            case "D26":
                // Len = F + M + N
                segs[0] = $"={sMid}";
                segs[7] = $"={sL}";
                segs[8] = $"={sR}";
                break;

            case "D06":
            case "D11":
                // Len = F + G (handles Left-hook OR Right-hook seamlessly)
                segs[0] = $"=IF({sL}>0,{sL},{sMid})";
                segs[1] = $"=IF({sL}>0,{sMid},{sR})";
                break;

            case "D03":
            case "D05":
            case "D12":
            case "D25":
            case "D27":
                // Len = F + G + M
                segs[0] = $"={sL}";
                segs[1] = $"={sMid}";
                segs[7] = $"={sR}";
                break;

            case "D07":
            case "D09":
            case "D10":
            case "D14":
            case "D15":
            case "D16":
            case "D34":
                // Len = F + G + H
                segs[0] = $"={sL}";
                segs[1] = $"={sMid}";
                segs[2] = $"={sR}";
                break;

            case "D13":
                // Len = F + G + M + N
                segs[0] = $"={sL1a}";
                segs[1] = $"={sMid}";
                segs[7] = $"={sL}";
                segs[8] = $"={sR}";
                break;

            case "D17":
            case "D20":
            case "D21":
            case "D30":
            case "D35":
                // Len = F + G + H + I
                segs[0] = $"={sL}";
                segs[1] = $"=IF({sL1a}<>\"\",{sL1a},{sMid})";
                segs[2] = $"=IF({sL1a}<>\"\",{sMid},{sR2a})";
                segs[3] = $"={sR}";
                break;

            case "D08":
            case "D18":
            case "D33":
            case "D36":
            case "D37":
            case "D38":
                // Len = F + G + H + I + J
                segs[0] = $"={sL}";
                segs[1] = $"={sL1a}";
                segs[2] = $"={sMid}";
                segs[3] = $"={sR2a}";
                segs[4] = $"={sR}";
                break;

            case "D19":
            case "D22":
            case "D23":
            case "D28":
            case "D29":
                // Len = F + G + H + I + J + K
                segs[0] = $"={sL}";
                segs[1] = $"={sL1a}";
                segs[2] = $"={sL1b}";
                segs[3] = $"={sMid}";
                segs[4] = $"={sR2a}";
                segs[5] = $"={sR}";
                break;

            case "D39":
                // Len = 2*(F + G) + 2*M
                segs[0] = $"=IF({q}V{rc}<>\"\",{sMid}/2,IF({q}N{rc}=\"PERIM_YH\",MAX(40,{q}G{rc}-2*{q}Q{rc}),IF({q}N{rc}=\"PERIM_XY_INNER\",MAX(40,ROUND(({q}F{rc}-2*{q}Q{rc})*0.67,0)),MAX(40,{q}F{rc}-2*{q}Q{rc}))))";
                segs[1] = $"=IF({q}V{rc}<>\"\",0,IF({q}N{rc}=\"PERIM_YH\",MAX(40,{q}H{rc}-2*{q}Q{rc}),MAX(40,{q}G{rc}-2*{q}Q{rc})))";
                segs[7] = $"={sL}";
                break;

            case "D41":
                // Len = 2*PI()*F + M
                segs[0] = $"=ROUND({sMid}/(2*PI()),1)";
                segs[7] = $"={sL}";
                break;

            default:
                // D31, D32, D40 (all 7 segments)
                segs[0] = $"={sL}";
                segs[1] = $"={sL1a}";
                segs[2] = $"={sL1b}";
                segs[3] = $"={sMid}";
                segs[4] = $"={sR2b}";
                segs[5] = $"={sR2a}";
                segs[6] = $"={sR}";
                break;
        }
        return segs;
    }

    private static object?[] ResolveFallbackSegments(int r, int d, string coverName, RebarConfigRowDto cfg, BarShapeMeta meta)
    {
        var segs = new object?[9];
        string cExpr = $"MAX({coverName},O{r})";
        string lMm = $"QS_Data!G{d}*1000";
        string wMm = $"QS_Data!H{d}*1000";
        string hMm = $"QS_Data!I{d}*1000";
        string bdir = !string.IsNullOrWhiteSpace(cfg.BarDir) ? cfg.BarDir.Trim().ToUpperInvariant() : (cfg.LenMode ?? "X").Trim().ToUpperInvariant();

        if (!string.IsNullOrWhiteSpace(cfg.OverrideLenExpr) && double.TryParse(cfg.OverrideLenExpr, NumberStyles.Any, CultureInfo.InvariantCulture, out double ovL))
        {
            segs[0] = ovL;
            return segs;
        }
        if (cfg.CustomL1Mm > 0)
        {
            segs[0] = cfg.CustomL1Mm;
            if (cfg.CustomL2Mm > 0) segs[1] = cfg.CustomL2Mm;
            if (cfg.CustomL3Mm > 0) segs[2] = cfg.CustomL3Mm;
            return segs;
        }

        string? midExpr = bdir switch
        {
            "Y" or "BY_W" => $"={wMm}-2*{cExpr}",
            "H" or "BY_H" => $"={hMm}-2*{cExpr}",
            "X*0.25" => $"=ROUND({lMm}*0.25,0)+SPEC_Lan_Tension*O{r}",
            "Y*0.25" => $"=ROUND({wMm}*0.25,0)+SPEC_Lan_Tension*O{r}",
            "X*0.7" => $"=ROUND({lMm}*0.7,0)",
            "SLOPED_XH" => $"=ROUND(SQRT(({lMm})^2+({hMm})^2),0)",
            "CHAIR_H" => $"=MAX(120,({hMm}-2*{cExpr})*2+150)",
            "TIE_Y" => $"=MAX(40,{wMm}-2*{cExpr})",
            "TIE_X" => $"=MAX(40,{lMm}-2*{cExpr})",
            _ => $"={lMm}-2*{cExpr}"
        };

        string hookL = ResolveHookExpr(cfg.HookLExpr, r);
        string hookR = ResolveHookExpr(cfg.HookRExpr, r);

        if (meta.Code == "D39")
        {
            if (bdir == "PERIM_XY_INNER")
            {
                segs[0] = $"=MAX(50,({lMm}-2*{cExpr})*0.67)";
                segs[1] = $"=MAX(50,{wMm}-2*{cExpr})";
            }
            else if (bdir == "PERIM_XY" || bdir == "STIRRUP_LW")
            {
                segs[0] = $"=MAX(50,{lMm}-2*{cExpr})";
                segs[1] = $"=MAX(50,{wMm}-2*{cExpr})";
            }
            else
            {
                segs[0] = $"=MAX(50,{wMm}-2*{cExpr})";
                segs[1] = $"=MAX(50,{hMm}-2*{cExpr})";
            }
            segs[7] = string.IsNullOrEmpty(hookL) ? $"=MAX(80,10*O{r})" : $"={hookL}";
        }
        else if (meta.Code is "D24" or "D02" or "D04")
        {
            segs[0] = midExpr;
            segs[7] = string.IsNullOrEmpty(hookL) ? $"=MAX(80,10*O{r})" : $"={hookL}";
            segs[8] = string.IsNullOrEmpty(hookR) ? segs[7] : $"={hookR}";
        }
        else if (meta.Code is "D06" or "D11")
        {
            segs[0] = string.IsNullOrEmpty(hookL) ? $"=MAX(250,SPEC_Lan_Tension*O{r})" : $"={hookL}";
            string extra = !string.IsNullOrEmpty(cfg.DeltaLenExpr) && bdir == "H" ? "+SPEC_Llap_Mult*O" + r : "";
            segs[1] = midExpr + extra;
        }
        else if (meta.Code is "D07" or "D15")
        {
            segs[0] = string.IsNullOrEmpty(hookL) ? $"=MAX(100,{hMm}-2*{cExpr})" : $"={hookL}";
            segs[1] = midExpr;
            segs[2] = string.IsNullOrEmpty(hookR) ? segs[0] : $"={hookR}";
        }
        else
        {
            segs[0] = midExpr;
        }
        return segs;
    }

    private static string ResolveHookExpr(string expr, int r)
    {
        if (string.IsNullOrWhiteSpace(expr)) return "";
        if (double.TryParse(expr.Trim(), NumberStyles.Any, CultureInfo.InvariantCulture, out double v))
        {
            if (v <= 0) return "";
            return v < 50 ? $"{Fmt(v)}*O{r}" : Fmt(v);
        }
        return "";
    }

    private static object ResolveFallbackQtyPerEl(int r, int d, string coverName, RebarConfigRowDto cfg)
    {
        double mult = cfg.LayerMult > 0 ? cfg.LayerMult : 1;
        if (!string.IsNullOrWhiteSpace(cfg.OverrideQtyExpr) && double.TryParse(cfg.OverrideQtyExpr, NumberStyles.Any, CultureInfo.InvariantCulture, out double qOv))
            return $"={Fmt(mult)}*{Fmt(qOv)}";

        string ddir = !string.IsNullOrWhiteSpace(cfg.DistDir) ? cfg.DistDir.Trim().ToUpperInvariant() : (cfg.DistMode ?? "FIXED").Trim().ToUpperInvariant();
        string lMm = $"QS_Data!G{d}*1000";
        string wMm = $"QS_Data!H{d}*1000";
        string hMm = $"QS_Data!I{d}*1000";
        double s1 = cfg.Space1Mm ?? ParseNumOr(cfg.CountOrSpace, 200);
        double s2 = cfg.Space2Mm ?? s1;

        return ddir switch
        {
            "X" or "SPREAD_L" => $"={Fmt(mult)}*(ROUND({lMm}/{Fmt(s1)},0)+1)",
            "Y" or "SPREAD_W" => $"={Fmt(mult)}*(ROUND({wMm}/{Fmt(s1)},0)+1)",
            "H" or "SPREAD_H" => $"={Fmt(mult)}*(ROUND({hMm}/{Fmt(s1)},0)+1)",
            "ZONE2_X" => $"={Fmt(mult)}*(ROUND(({lMm}*0.5)/{Fmt(s1)},0)+ROUND(({lMm}*0.5)/{Fmt(s2)},0)+2)",
            "ZONE2_H" => $"={Fmt(mult)}*(ROUND(({hMm}*0.5)/{Fmt(s1)},0)+ROUND(({hMm}*0.5)/{Fmt(s2)},0)+1)",
            "XY_GRID" => $"={Fmt(mult)}*(ROUND({lMm}/{Fmt(s1)},0)+1)*(ROUND({wMm}/{Fmt(s2)},0)+1)",
            "XH_GRID" => $"={Fmt(mult)}*(ROUND({lMm}/{Fmt(s1)},0)+1)*(ROUND({hMm}/{Fmt(s2)},0)+1)",
            _ => $"={Fmt(mult)}*{Fmt(ParseNumOr(cfg.CountOrSpace, 1))}"
        };
    }

    private SheetCellRowDto BuildManualTktRow(int r, int stt, ExistingTktRowDto m)
    {
        var meta = BarShapeCatalog.Get(m.ShapeCode);
        string lenExpr = meta.LenExpr.Replace("{r}", r.ToString(CultureInfo.InvariantCulture));
        string geomLen = meta.BendCount > 0
            ? $"(({lenExpr})-{meta.BendCount}*IFERROR(SPEC_BendDeduct90,IFERROR(CFG_BendDeduct,1))*O{r})"
            : $"({lenExpr})";

        string barLenFormula = meta.HasLap
            ? $"=IF(O{r}=\"\",\"\",ROUND(({geomLen}+IF(AND(O{r}>=10,O{r}<IFERROR(SPEC_CouplerMinD,IFERROR(CFG_CouplerMinD,20))),MAX(0,CEILING({geomLen}/IFERROR(SPEC_StdBarLen,IFERROR(CFG_BarStdLen,11700)),1)-1)*IFERROR(SPEC_Llap_Mult,IFERROR(CFG_LapRatio,45))*O{r},0))/1000,3))"
            : $"=IF(O{r}=\"\",\"\",ROUND({geomLen}/1000,3))";

        var vals = new object?[TotalCols];
        vals[0] = stt;
        vals[1] = m.CompName;
        vals[2] = m.BarMark;
        vals[3] = meta.Code;
        vals[4] = meta.Code;
        vals[5] = NullIfEmpty(m.L1Expr); vals[6] = NullIfEmpty(m.L2Expr); vals[7] = NullIfEmpty(m.L3Expr);
        vals[8] = NullIfEmpty(m.L4Expr); vals[9] = NullIfEmpty(m.L5Expr); vals[10] = NullIfEmpty(m.L6Expr); vals[11] = NullIfEmpty(m.L7Expr);
        vals[12] = NullIfEmpty(m.M1Expr); vals[13] = NullIfEmpty(m.M2Expr);
        vals[14] = m.Dia > 0 ? m.Dia : 12;
        vals[15] = NullIfEmpty(m.QtyPerElExpr) ?? 1;
        vals[16] = NullIfEmpty(m.ElQtyExpr) ?? 1;
        vals[17] = $"=IF(O{r}=\"\",\"\",P{r}*Q{r})";
        vals[18] = barLenFormula;
        vals[19] = meta.HasLap ? $"=IF(OR(O{r}=\"\",S{r}=\"\"),0,IF(O{r}>=IFERROR(SPEC_CouplerMinD,IFERROR(CFG_CouplerMinD,20)),R{r}*MAX(0,CEILING({geomLen}/IFERROR(SPEC_StdBarLen,IFERROR(CFG_BarStdLen,11700)),1)-1),0))" : 0;
        vals[20] = $"=IF(O{r}=\"\",\"\",ROUND(R{r}*S{r},2))";
        vals[21] = $"=IF(O{r}=\"\",\"\",ROUND(PI()*(O{r}/1000)^2/4*7850,3))";
        vals[22] = $"=IF(O{r}=\"\",\"\",ROUND(U{r}*V{r},2))";
        vals[23] = m.Note;
        vals[24] = "MANUAL";
        vals[25] = string.IsNullOrWhiteSpace(m.BoqBucket) ? "Misc/Stair/Lintel" : m.BoqBucket;

        return new SheetCellRowDto
        {
            RowKind = "MANUAL",
            OutlineLevel = 0,
            TagId = "MANUAL",
            SectionKey = vals[25]?.ToString() ?? "",
            Values = vals.ToList()
        };
    }

    private static object? NullIfEmpty(string s) => string.IsNullOrWhiteSpace(s) ? null : s;

    private static string ResolveCoverName(string? group) => (group ?? "").Trim().ToLowerInvariant() switch
    {
        "foundation" or "groundbeam" or "slabongrade" => "SPEC_Cover_Ftg",
        "column" => "SPEC_Cover_Col",
        "beam" or "lintel" => "SPEC_Cover_Beam",
        "wall" or "tank" or "manhole" => "SPEC_Cover_Wall",
        _ => "SPEC_Cover_Slab"
    };

    private static bool IsFixedSpec(string s) =>
        !string.IsNullOrWhiteSpace(s) && !s.TrimStart().StartsWith('@') && !s.Contains('/');

    private static (double s1, double s2) ParseDualSpacing(string s, double def1, double def2)
    {
        if (string.IsNullOrWhiteSpace(s)) return (def1, def2);
        var parts = s.Split('/');
        if (parts.Length >= 2 &&
            double.TryParse(parts[0].Trim().TrimStart('@'), NumberStyles.Any, CultureInfo.InvariantCulture, out double a) && a > 0 &&
            double.TryParse(parts[1].Trim().TrimStart('@'), NumberStyles.Any, CultureInfo.InvariantCulture, out double b) && b > 0)
            return (a, b);
        double single = ParseNumOr(s, def1);
        return (single, single);
    }

    private static double ParseNumOr(string s, double def) =>
        double.TryParse((s ?? "").Trim().TrimStart('@'), NumberStyles.Any, CultureInfo.InvariantCulture, out double v) && v > 0 ? v : def;

    private static string Fmt(double v) => v.ToString("0.###", CultureInfo.InvariantCulture);
}
