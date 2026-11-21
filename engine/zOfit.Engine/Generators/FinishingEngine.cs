using System.Globalization;
using ZOfit.Engine.Models;

namespace ZOfit.Engine.Generators;

public sealed class FinishingEngine
{
    public GeneratedSheetDto Generate(GenerateFinishingRequest req)
    {
        var rows = new List<SheetCellRowDto>();
        rows.Add(RcDetailGenerator.MakeRow("TITLE", 0, "HDR_MARKER", "HDR", "ZOFIT_FINISHING_V2", "FINISHING WORKS DETAIL TAKEOFF / BÓC TÁCH CHI TIẾT HOÀN THIỆN", req.Spec.ProjectName, "", "", "", "", "", "", "", "", "", "", ""));
        rows.Add(RcDetailGenerator.MakeRow("BANNER", 0, "HDR_PROJ", "HDR", "", "Project:", $"=IFERROR(CFG_ProjectName,\"{req.Spec.ProjectName.Replace("\"", "\"\"")}\")", "", "Rule:", "Opening > SPEC_Wall_OpeningMin (INDEX/MATCH QS_DoorWin)", "", "", "", "", "", "", "HDR_PROJ", "HDR"));
        rows.Add(RcDetailGenerator.MakeRow("BLANK", 0, "", "HDR"));
        rows.Add(RcDetailGenerator.MakeRow("WORK_HDR", 0, "HDR_COLS", "HDR", "Code", "Description", "Spec / Dim", "Unit", "NO", "L (m)", "W (m)", "H / t (m)", "a", "b", "c", "Total Q'TY", "TagID (Sync)", "SectionKey"));
        rows.Add(RcDetailGenerator.MakeRow("BLANK", 0, "", "HDR"));

        var manualBySection = GroupManual(req.ExistingRows);
        var rooms = req.Rooms ?? new List<RoomFinishDto>();

        BuildRoofAndWetSection(rows, rooms, manualBySection);
        BuildMasonryAndWallSection(rows, rooms, manualBySection);
        BuildFloorAndCeilingSection(rows, rooms, manualBySection);

        int manualPreserved = manualBySection.Values.Sum(v => v.Count);
        return new GeneratedSheetDto
        {
            SheetName = "Finishing_Detail",
            Building = "FINISHING",
            AddedCount = rooms.Count,
            UpdatedCount = 0,
            PreservedManualCount = manualPreserved,
            RemovedFlaggedCount = 0,
            Rows = rows
        };
    }

    private void BuildRoofAndWetSection(List<SheetCellRowDto> rows, List<RoomFinishDto> rooms, Dictionary<string, List<ExistingDetailRowDto>> manualMap)
    {
        var roofRooms = rooms.Where(r => r.RoofAreaM2 > 0).ToList();
        var wetRooms = rooms.Where(r => string.Equals(r.IsWetArea?.Trim(), "Y", StringComparison.OrdinalIgnoreCase)).ToList();
        if (roofRooms.Count == 0 && wetRooms.Count == 0 && !manualMap.ContainsKey("FIN_ROOF")) return;

        rows.Add(RcDetailGenerator.MakeRow("WORK_HDR", 0, "SEC_ROOF", "FIN_ROOF", "I", "ROOFING & WATERPROOFING WORKS / CÔNG TÁC MÁI & CHỐNG THẤM", "", "", null, null, null, null, null, null, null, null, "SEC_ROOF", "FIN_ROOF"));

        if (roofRooms.Count > 0)
        {
            // Base roof waterproofing 2mm (chain anchor)
            int hdrWp2 = NextRow(rows);
            rows.Add(RcDetailGenerator.MakeRow("ITEM_HDR", 1, "FIN_ROOF_WP2", "FIN_ROOF", "1", "Exposed Urethane Waterproofing 2mm (UV) / Chống thấm mái 2mm", "t=2mm UV", "m2", null, null, null, null, null, null, null, 0, "FIN_ROOF_WP2", "FIN_ROOF"));
            int s = NextRow(rows);
            foreach (var rm in roofRooms)
            {
                int r = NextRow(rows);
                int dr = rm.DataRow;
                rows.Add(RcDetailGenerator.MakeRow("DATA", 2, $"{rm.RoomId}::ROOF_WP2", "FIN_ROOF",
                    rm.RoomId, $"Roof {rm.RoomName} ({rm.Building})", "", "m2",
                    1, $"=QS_Rooms!N{dr}", null, null, null, null, null, $"=PRODUCT(E{r}:K{r})", $"{rm.RoomId}::ROOF_WP2", "FIN_ROOF"));
            }
            int e = NextRow(rows) - 1;
            rows[hdrWp2 - 1].Values[11] = $"=SUBTOTAL(9,L{s}:L{e})";

            // Chained Roof Layers linked directly to L{hdrWp2}!
            int rSlope = NextRow(rows);
            rows.Add(RcDetailGenerator.MakeRow("ITEM_HDR", 1, "FIN_ROOF_SLOPE", "FIN_ROOF", "2", "Slope screed plain concrete M250 / BT tạo dốc mái M250", "t=100..180mm", "m3",
                1, $"=L{hdrWp2}", null, 0.15, "=SPEC_Conc_Waste", null, null, $"=PRODUCT(E{rSlope}:K{rSlope})", "FIN_ROOF_SLOPE", "FIN_ROOF"));

            int rMesh = NextRow(rows);
            rows.Add(RcDetailGenerator.MakeRow("ITEM_HDR", 1, "FIN_ROOF_MESH", "FIN_ROOF", "3", "Roof wire mesh #8-150x150 / Lưới thép hàn mái", "#8-150x150", "m2",
                1, $"=L{hdrWp2}", null, null, null, null, null, $"=PRODUCT(E{rMesh}:K{rMesh})", "FIN_ROOF_MESH", "FIN_ROOF"));

            int rWp3 = NextRow(rows);
            rows.Add(RcDetailGenerator.MakeRow("ITEM_HDR", 1, "FIN_ROOF_WP3", "FIN_ROOF", "4", "Non-exposed Urethane Waterproofing 3mm / Chống thấm 3mm", "t=3mm", "m2",
                1, $"=L{hdrWp2}", null, null, null, null, null, $"=PRODUCT(E{rWp3}:K{rWp3})", "FIN_ROOF_WP3", "FIN_ROOF"));

            int rPol = NextRow(rows);
            rows.Add(RcDetailGenerator.MakeRow("ITEM_HDR", 1, "FIN_ROOF_POL", "FIN_ROOF", "5", "Roof polishing & power float / Đánh bóng xoa mặt mái", "", "m2",
                1, $"=L{hdrWp2}", null, null, null, null, null, $"=PRODUCT(E{rPol}:K{rPol})", "FIN_ROOF_POL", "FIN_ROOF"));

            // Parapet masonry if ParapetLenM > 0
            var parapets = roofRooms.Where(x => x.ParapetLenM > 0).ToList();
            if (parapets.Count > 0)
            {
                int hdrPar = NextRow(rows);
                rows.Add(RcDetailGenerator.MakeRow("ITEM_HDR", 1, "FIN_PARAPET", "FIN_ROOF", "6", "Parapet brick wall / Xây gạch tường chắn mái Parapet", "t=150mm", "m3", null, null, null, null, null, null, null, 0, "FIN_PARAPET", "FIN_ROOF"));
                int ps = NextRow(rows);
                foreach (var pm in parapets)
                {
                    int r = NextRow(rows);
                    int dr = pm.DataRow;
                    rows.Add(RcDetailGenerator.MakeRow("DATA", 2, $"{pm.RoomId}::PARAPET", "FIN_ROOF",
                        pm.RoomId, "Parapet " + pm.RoomName, "", "m3",
                        1, $"=QS_Rooms!O{dr}", 0.15, $"=QS_Rooms!P{dr}", null, null, null, $"=PRODUCT(E{r}:K{r})", $"{pm.RoomId}::PARAPET", "FIN_ROOF"));
                }
                int pe = NextRow(rows) - 1;
                rows[hdrPar - 1].Values[11] = $"=SUBTOTAL(9,L{ps}:L{pe})";
            }
        }

        if (wetRooms.Count > 0)
        {
            int hdrWet = NextRow(rows);
            rows.Add(RcDetailGenerator.MakeRow("ITEM_HDR", 1, "FIN_WET_WP", "FIN_ROOF", "7", "Wet area waterproofing (+300mm curb) / Chống thấm WC (+vén chân tường)", "SPEC_WP_CurbH", "m2", null, null, null, null, null, null, null, 0, "FIN_WET_WP", "FIN_ROOF"));
            int wsStart = NextRow(rows);
            foreach (var wm in wetRooms)
            {
                int r = NextRow(rows);
                int dr = wm.DataRow;
                rows.Add(RcDetailGenerator.MakeRow("DATA", 2, $"{wm.RoomId}::WET_WP", "FIN_ROOF",
                    wm.RoomId, $"WP {wm.RoomName} (Floor + Curb)", "", "m2",
                    1, $"=QS_Rooms!E{dr}", $"=QS_Rooms!F{dr}", "=SPEC_WP_CurbH", null, null, null,
                    $"=F{r}+G{r}*H{r}", $"{wm.RoomId}::WET_WP", "FIN_ROOF"));
            }
            int wsEnd = NextRow(rows) - 1;
            rows[hdrWet - 1].Values[11] = $"=SUBTOTAL(9,L{wsStart}:L{wsEnd})";
        }
        AppendManual(rows, manualMap, "FIN_ROOF");
        rows.Add(RcDetailGenerator.MakeRow("BLANK", 0, "", "FIN_ROOF"));
    }

    private void BuildMasonryAndWallSection(List<SheetCellRowDto> rows, List<RoomFinishDto> rooms, Dictionary<string, List<ExistingDetailRowDto>> manualMap)
    {
        var wallRooms = rooms.Where(r => r.WallLenM > 0 && !string.Equals(r.WallType, "NONE", StringComparison.OrdinalIgnoreCase)).ToList();
        if (wallRooms.Count == 0 && !manualMap.ContainsKey("FIN_WALL")) return;

        rows.Add(RcDetailGenerator.MakeRow("WORK_HDR", 0, "SEC_WALL", "FIN_WALL", "II", "MASONRY, PLASTER, TILING & PAINT (AUTO-DEDUCT DOORS) / XÂY, TRÁT, ỐP, SƠN", "", "", null, null, null, null, null, null, null, null, "SEC_WALL", "FIN_WALL"));

        // 1. Brick Masonry (m3)
        var brickRooms = wallRooms.Where(r => !string.Equals(r.WallType, "PANEL", StringComparison.OrdinalIgnoreCase)).ToList();
        if (brickRooms.Count > 0)
        {
            int hdrBr = NextRow(rows);
            rows.Add(RcDetailGenerator.MakeRow("ITEM_HDR", 1, "FIN_MASONRY", "FIN_WALL", "1", "Brick masonry wall (deduct openings >0.25m2 via QS_DoorWin) / Xây tường gạch", "", "m3", null, null, null, null, null, null, null, 0, "FIN_MASONRY", "FIN_WALL"));
            int bs = NextRow(rows);
            foreach (var rm in brickRooms)
            {
                int r = NextRow(rows);
                int dr = rm.DataRow;
                double thk = ResolveWallThicknessM(rm.WallType);
                string doorAreaExpr = BuildDoorLookupExpr(rm.DoorCodes, true, "SPEC_Wall_OpeningMin");
                rows.Add(RcDetailGenerator.MakeRow("DATA", 2, $"{rm.RoomId}::MASONRY", "FIN_WALL",
                    rm.RoomId, $"Brick wall {rm.RoomName} ({rm.WallType})", rm.DoorCodes, "m3",
                    1, $"=QS_Rooms!I{dr}", thk, $"=QS_Rooms!G{dr}", null, null, null,
                    $"=MAX(0,(F{r}*H{r}-({doorAreaExpr}))*G{r})", $"{rm.RoomId}::MASONRY", "FIN_WALL"));
            }
            int be = NextRow(rows) - 1;
            rows[hdrBr - 1].Values[11] = $"=SUBTOTAL(9,L{bs}:L{be})";

            // 2. Wall Plastering (2 faces - 2*door + jamb)
            int hdrPl = NextRow(rows);
            rows.Add(RcDetailGenerator.MakeRow("ITEM_HDR", 1, "FIN_PLASTER", "FIN_WALL", "2", "Internal/External wall plastering 2 faces / Trát tường (trừ cửa)", "", "m2", null, null, null, null, null, null, null, 0, "FIN_PLASTER", "FIN_WALL"));
            int pls = NextRow(rows);
            foreach (var rm in brickRooms)
            {
                int r = NextRow(rows);
                int dr = rm.DataRow;
                string doorAreaExpr = BuildDoorLookupExpr(rm.DoorCodes, true, "SPEC_Wall_OpeningMin");
                rows.Add(RcDetailGenerator.MakeRow("DATA", 2, $"{rm.RoomId}::PLASTER", "FIN_WALL",
                    rm.RoomId, $"Plaster {rm.RoomName}", rm.DoorCodes, "m2",
                    2, $"=QS_Rooms!I{dr}", $"=QS_Rooms!G{dr}", null, null, null, null,
                    $"=MAX(0,E{r}*(F{r}*G{r}-({doorAreaExpr})))", $"{rm.RoomId}::PLASTER", "FIN_WALL"));
            }
            int ple = NextRow(rows) - 1;
            rows[hdrPl - 1].Values[11] = $"=SUBTOTAL(9,L{pls}:L{ple})";

            // 3. Anti-crack wire mesh (width SPEC_Mesh_Width along columns & beam soffit, 2 faces)
            int hdrMesh = NextRow(rows);
            rows.Add(RcDetailGenerator.MakeRow("ITEM_HDR", 1, "FIN_CRACKMESH", "FIN_WALL", "3", "Anti-crack wire mesh at concrete-brick interface / Lưới mắt cáo chống nứt", "W=SPEC_Mesh_Width", "m2", null, null, null, null, null, null, null, 0, "FIN_CRACKMESH", "FIN_WALL"));
            int ms = NextRow(rows);
            foreach (var rm in brickRooms)
            {
                int r = NextRow(rows);
                int dr = rm.DataRow;
                rows.Add(RcDetailGenerator.MakeRow("DATA", 2, $"{rm.RoomId}::MESH", "FIN_WALL",
                    rm.RoomId, $"Crack mesh {rm.RoomName}", "", "m2",
                    2, $"=2*QS_Rooms!G{dr}+QS_Rooms!I{dr}", "=SPEC_Mesh_Width", null, null, null, null,
                    $"=PRODUCT(E{r}:K{r})", $"{rm.RoomId}::MESH", "FIN_WALL"));
            }
            int me = NextRow(rows) - 1;
            rows[hdrMesh - 1].Values[11] = $"=SUBTOTAL(9,L{ms}:L{me})";

            // 4. Wall Tiling (m2) & Wall Paint (m2 = Plaster - Tiling)
            int hdrTile = NextRow(rows);
            rows.Add(RcDetailGenerator.MakeRow("ITEM_HDR", 1, "FIN_WALLTILE", "FIN_WALL", "4", "Wall ceramic tiling / Ốp gạch tường", "", "m2", null, null, null, null, null, null, null, 0, "FIN_WALLTILE", "FIN_WALL"));
            int ts = NextRow(rows);
            foreach (var rm in brickRooms.Where(x => x.WallTileHeightM > 0))
            {
                int r = NextRow(rows);
                int dr = rm.DataRow;
                rows.Add(RcDetailGenerator.MakeRow("DATA", 2, $"{rm.RoomId}::WTILE", "FIN_WALL",
                    rm.RoomId, $"Wall tile {rm.RoomName}", "", "m2",
                    1, $"=QS_Rooms!I{dr}", $"=QS_Rooms!L{dr}", null, null, null, null,
                    $"=PRODUCT(E{r}:K{r})", $"{rm.RoomId}::WTILE", "FIN_WALL"));
            }
            int te = NextRow(rows) - 1;
            rows[hdrTile - 1].Values[11] = te >= ts ? $"=SUBTOTAL(9,L{ts}:L{te})" : 0;

            int rPaint = NextRow(rows);
            rows.Add(RcDetailGenerator.MakeRow("ITEM_HDR", 1, "FIN_PAINT", "FIN_WALL", "5", "Emulsion paint (1 primer + 2 coats = Plaster - Wall Tiling) / Sơn tường", "TCVN 6934", "m2",
                null, null, null, null, null, null, null, $"=MAX(0,L{hdrPl}-L{hdrTile})", "FIN_PAINT", "FIN_WALL"));
        }

        // PU Panel partitions if any
        var panelRooms = wallRooms.Where(r => string.Equals(r.WallType, "PANEL", StringComparison.OrdinalIgnoreCase)).ToList();
        if (panelRooms.Count > 0)
        {
            int hdrPan = NextRow(rows);
            rows.Add(RcDetailGenerator.MakeRow("ITEM_HDR", 1, "FIN_PANEL", "FIN_WALL", "6", "PU insulated panel partition (deduct openings) / Vách ngăn Panel PU", "", "m2", null, null, null, null, null, null, null, 0, "FIN_PANEL", "FIN_WALL"));
            int ps = NextRow(rows);
            foreach (var rm in panelRooms)
            {
                int r = NextRow(rows);
                int dr = rm.DataRow;
                string doorAreaExpr = BuildDoorLookupExpr(rm.DoorCodes, false, "0");
                rows.Add(RcDetailGenerator.MakeRow("DATA", 2, $"{rm.RoomId}::PANEL", "FIN_WALL",
                    rm.RoomId, $"Panel {rm.RoomName}", rm.DoorCodes, "m2",
                    1, $"=QS_Rooms!I{dr}", $"=QS_Rooms!G{dr}", null, null, null, null,
                    $"=MAX(0,F{r}*G{r}-({doorAreaExpr}))", $"{rm.RoomId}::PANEL", "FIN_WALL"));
            }
            int pe = NextRow(rows) - 1;
            rows[hdrPan - 1].Values[11] = $"=SUBTOTAL(9,L{ps}:L{pe})";
        }

        AppendManual(rows, manualMap, "FIN_WALL");
        rows.Add(RcDetailGenerator.MakeRow("BLANK", 0, "", "FIN_WALL"));
    }

    private void BuildFloorAndCeilingSection(List<SheetCellRowDto> rows, List<RoomFinishDto> rooms, Dictionary<string, List<ExistingDetailRowDto>> manualMap)
    {
        var floorRooms = rooms.Where(r => r.AreaM2 > 0).ToList();
        if (floorRooms.Count == 0 && !manualMap.ContainsKey("FIN_FLOOR")) return;

        rows.Add(RcDetailGenerator.MakeRow("WORK_HDR", 0, "SEC_FLOOR", "FIN_FLOOR", "III", "FLOORING, SKIRTING & CEILING WORKS / LÁT NỀN, LEN CHÂN TƯỜNG & TRẦN", "", "", null, null, null, null, null, null, null, null, "SEC_FLOOR", "FIN_FLOOR"));

        // 1. Floor Finish
        int hdrFl = NextRow(rows);
        rows.Add(RcDetailGenerator.MakeRow("ITEM_HDR", 1, "FIN_FLOOR_FIN", "FIN_FLOOR", "1", "Floor tiling / Epoxy / Hardener finish / Hoàn thiện lát nền", "", "m2", null, null, null, null, null, null, null, 0, "FIN_FLOOR_FIN", "FIN_FLOOR"));
        int fs = NextRow(rows);
        foreach (var rm in floorRooms)
        {
            int r = NextRow(rows);
            int dr = rm.DataRow;
            rows.Add(RcDetailGenerator.MakeRow("DATA", 2, $"{rm.RoomId}::FLOOR", "FIN_FLOOR",
                rm.RoomId, $"{rm.FloorFinish} - {rm.RoomName}", rm.FloorFinish, "m2",
                1, $"=QS_Rooms!E{dr}", null, null, null, null, null, $"=PRODUCT(E{r}:K{r})", $"{rm.RoomId}::FLOOR", "FIN_FLOOR"));
        }
        int fe = NextRow(rows) - 1;
        rows[hdrFl - 1].Values[11] = $"=SUBTOTAL(9,L{fs}:L{fe})";

        // 2. Skirting (Perimeter - door widths)
        int hdrSk = NextRow(rows);
        rows.Add(RcDetailGenerator.MakeRow("ITEM_HDR", 1, "FIN_SKIRT", "FIN_FLOOR", "2", "Skirting (Perimeter - Door widths) / Len chân tường", "", "m", null, null, null, null, null, null, null, 0, "FIN_SKIRT", "FIN_FLOOR"));
        int ss = NextRow(rows);
        foreach (var rm in floorRooms.Where(x => x.PerimM > 0))
        {
            int r = NextRow(rows);
            int dr = rm.DataRow;
            string doorWExpr = BuildDoorWidthLookupExpr(rm.DoorCodes);
            rows.Add(RcDetailGenerator.MakeRow("DATA", 2, $"{rm.RoomId}::SKIRT", "FIN_FLOOR",
                rm.RoomId, $"Skirting - {rm.RoomName}", rm.DoorCodes, "m",
                1, $"=QS_Rooms!F{dr}", null, null, null, null, null,
                $"=MAX(0,F{r}-({doorWExpr}))", $"{rm.RoomId}::SKIRT", "FIN_FLOOR"));
        }
        int se = NextRow(rows) - 1;
        rows[hdrSk - 1].Values[11] = se >= ss ? $"=SUBTOTAL(9,L{ss}:L{se})" : 0;

        // 3. False Ceiling
        var ceilRooms = floorRooms.Where(x => !string.Equals(x.CeilingType, "NONE", StringComparison.OrdinalIgnoreCase)).ToList();
        if (ceilRooms.Count > 0)
        {
            int hdrCl = NextRow(rows);
            rows.Add(RcDetailGenerator.MakeRow("ITEM_HDR", 1, "FIN_CEILING", "FIN_FLOOR", "3", "Gypsum ceiling / Direct ceiling paint / Trần thạch cao & Sơn trần", "", "m2", null, null, null, null, null, null, null, 0, "FIN_CEILING", "FIN_FLOOR"));
            int cs = NextRow(rows);
            foreach (var rm in ceilRooms)
            {
                int r = NextRow(rows);
                int dr = rm.DataRow;
                rows.Add(RcDetailGenerator.MakeRow("DATA", 2, $"{rm.RoomId}::CEIL", "FIN_FLOOR",
                    rm.RoomId, $"{rm.CeilingType} Ceiling - {rm.RoomName}", rm.CeilingType, "m2",
                    1, $"=QS_Rooms!E{dr}", null, null, null, null, null, $"=PRODUCT(E{r}:K{r})", $"{rm.RoomId}::CEIL", "FIN_FLOOR"));
            }
            int ce = NextRow(rows) - 1;
            rows[hdrCl - 1].Values[11] = $"=SUBTOTAL(9,L{cs}:L{ce})";
        }

        AppendManual(rows, manualMap, "FIN_FLOOR");
    }

    public static string BuildDoorLookupExpr(string? doorCodes, bool applyThreshold, string thresholdName)
    {
        if (string.IsNullOrWhiteSpace(doorCodes)) return "0";
        var parts = ParseDoorCodes(doorCodes);
        if (parts.Count == 0) return "0";
        var terms = parts.Select(p =>
        {
            string lookup = $"IFERROR(INDEX(QS_DoorWin!$E$5:$E$500,MATCH(\"{p.Code}\",QS_DoorWin!$A$5:$A$500,0)),0)";
            string unitArea = applyThreshold ? $"IF({lookup}>{thresholdName},{lookup},0)" : lookup;
            return p.Count == 1 ? unitArea : $"{p.Count}*{unitArea}";
        });
        return string.Join("+", terms);
    }

    public static string BuildDoorWidthLookupExpr(string? doorCodes)
    {
        if (string.IsNullOrWhiteSpace(doorCodes)) return "0";
        var parts = ParseDoorCodes(doorCodes);
        if (parts.Count == 0) return "0";
        var terms = parts.Select(p =>
        {
            string lookup = $"IFERROR(INDEX(QS_DoorWin!$C$5:$C$500,MATCH(\"{p.Code}\",QS_DoorWin!$A$5:$A$500,0)),0)";
            return p.Count == 1 ? lookup : $"{p.Count}*{lookup}";
        });
        return string.Join("+", terms);
    }

    private static List<(string Code, int Count)> ParseDoorCodes(string doorCodes)
    {
        var list = new List<(string, int)>();
        foreach (var raw in doorCodes.Split(new[] { ',', ';', '+' }, StringSplitOptions.RemoveEmptyEntries))
        {
            string tok = raw.Trim();
            if (string.IsNullOrEmpty(tok)) continue;
            int cnt = 1;
            string code = tok;
            int star = tok.IndexOf('*');
            if (star > 0)
            {
                code = tok.Substring(0, star).Trim();
                int.TryParse(tok.Substring(star + 1).Trim(), NumberStyles.Any, CultureInfo.InvariantCulture, out cnt);
                if (cnt <= 0) cnt = 1;
            }
            if (!string.IsNullOrEmpty(code)) list.Add((code, cnt));
        }
        return list;
    }

    private static double ResolveWallThicknessM(string? wallType)
    {
        string wt = (wallType ?? "").Trim().ToUpperInvariant();
        if (wt.Contains("100")) return 0.10;
        if (wt.Contains("110")) return 0.11;
        if (wt.Contains("150")) return 0.15;
        if (wt.Contains("220")) return 0.22;
        return 0.20;
    }

    private static Dictionary<string, List<ExistingDetailRowDto>> GroupManual(IReadOnlyList<ExistingDetailRowDto>? existing)
    {
        var dict = new Dictionary<string, List<ExistingDetailRowDto>>(StringComparer.OrdinalIgnoreCase);
        if (existing == null) return dict;
        foreach (var row in existing)
        {
            if (!string.Equals(row.TagId, "MANUAL", StringComparison.OrdinalIgnoreCase)) continue;
            string sec = string.IsNullOrWhiteSpace(row.SectionKey) ? "FIN_FLOOR" : row.SectionKey.Trim();
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
