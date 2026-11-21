using System.Text.RegularExpressions;
using ZOfit.Engine.Models;

namespace ZOfit.Engine.Parsers;

/// <summary>
/// Parses beam cross-section tables from AutoCAD (e.g., Image 0: 1MBX1 (800x800) with Left/Mid/Right sections).
/// Extracts beam dimensions, top/bottom bars, stirrups (A, B, C with zones), and side bars.
/// Converts directly into ElementDataDto (QS_Data) and RebarConfigRowDto (QS_RebarConfig).
/// </summary>
public sealed partial class CadBeamSectionParser
{
    [GeneratedRegex(@"(?<name>[0-9A-Za-z_-]+)\s*(?:\((?<b>\d+)[xX*](?<h>\d+)\)|(?<b>\d+)[xX*](?<h>\d+))", RegexOptions.IgnoreCase)]
    private static partial Regex BeamNameAndSizeRegex();

    [GeneratedRegex(@"(?<count>\d+)\s*(?:D|d|phi|%%c|-D|\s+D)\s*(?<dia>\d+)", RegexOptions.IgnoreCase)]
    private static partial Regex BarCountAndDiaRegex();

    [GeneratedRegex(@"(?:(?:A|B|C|D|B,C|B,C,D|BC|BCD)\s*:\s*)?(?:D|d|phi|%%c)?(?<dia>\d+)\s*a\s*(?<spacing>\d+)", RegexOptions.IgnoreCase)]
    private static partial Regex StirrupSpecRegex();

    [GeneratedRegex(@"(?:D|d|phi|%%c)?(?<dia>\d+)\s*a\s*(?<spacing>\d+)", RegexOptions.IgnoreCase)]
    private static partial Regex SupportBarRegex();

    [GeneratedRegex(@"FL\s*(?<sign>[+-]?)(?<val>[\d.]+)", RegexOptions.IgnoreCase)]
    private static partial Regex LevelRegex();

    public CadBeamSectionParseResponse Parse(CadBeamSectionParseRequest req)
    {
        var resp = new CadBeamSectionParseResponse();
        var rawEntities = req.Entities ?? new List<CadEntityItemDto>();

        // Normalize text in all entities using TCVN3 converter
        foreach (var e in rawEntities)
        {
            e.Text = Tcvn3Converter.ToUnicode(e.Text).Trim();
        }

        // 1. Detect Beam Names and Dimensions
        var detectedBeams = new List<(string name, double b, double h, double elev)>();

        // 1a. From Block References (e.g. HA-ElementName)
        foreach (var bRef in rawEntities.Where(e => e.Type.Equals("INSERT", StringComparison.OrdinalIgnoreCase) && e.Attributes != null))
        {
            var attr = bRef.Attributes!;
            string tenck = attr.GetValueOrDefault("TENCK", "").Trim();
            string kt = attr.GetValueOrDefault("KT", "").Trim();

            if (!string.IsNullOrEmpty(tenck))
            {
                var m = BeamNameAndSizeRegex().Match($"{tenck} {kt}");
                if (m.Success)
                {
                    double w = double.Parse(m.Groups["b"].Value);
                    double ht = double.Parse(m.Groups["h"].Value);
                    detectedBeams.Add((m.Groups["name"].Value.ToUpperInvariant(), w, ht, 0.0));
                }
            }
        }

        // 1b. From Text/MText scan if no block found
        if (detectedBeams.Count == 0)
        {
            foreach (var t in rawEntities.Where(e => !string.IsNullOrEmpty(e.Text)))
            {
                var matches = BeamNameAndSizeRegex().Matches(t.Text);
                foreach (Match m in matches)
                {
                    double w = double.Parse(m.Groups["b"].Value);
                    double ht = double.Parse(m.Groups["h"].Value);
                    string bName = m.Groups["name"].Value.ToUpperInvariant();
                    if (!detectedBeams.Any(x => x.name.Equals(bName, StringComparison.OrdinalIgnoreCase)))
                    {
                        detectedBeams.Add((bName, w, ht, 0.0));
                    }
                }
            }
        }

        // 1c. From raw text lines fallback
        if (detectedBeams.Count == 0 && req.RawTextLines != null)
        {
            foreach (var line in req.RawTextLines)
            {
                string norm = Tcvn3Converter.ToUnicode(line);
                var matches = BeamNameAndSizeRegex().Matches(norm);
                foreach (Match m in matches)
                {
                    double w = double.Parse(m.Groups["b"].Value);
                    double ht = double.Parse(m.Groups["h"].Value);
                    string bName = m.Groups["name"].Value.ToUpperInvariant();
                    if (!detectedBeams.Any(x => x.name.Equals(bName, StringComparison.OrdinalIgnoreCase)))
                    {
                        detectedBeams.Add((bName, w, ht, 0.0));
                    }
                }
            }
        }

        // Default fallback if still not found
        if (detectedBeams.Count == 0)
        {
            detectedBeams.Add(("1MBX1", 800, 800, 0.0));
            resp.Messages.Add("Không tìm thấy tên dầm trong bảng; dùng mặc định 1MBX1 (800x800).");
        }

        // 2. Locate Columns (TRÁI/LEFT, GIỮA/CEN, PHẢI/RIGHT)
        double xLeft = double.NaN, xMid = double.NaN, xRight = double.NaN;
        foreach (var t in rawEntities.Where(e => e.Position.Length >= 2))
        {
            string up = t.Text.ToUpperInvariant();
            if (up.Contains("TRÁI") || up.Contains("LEFT"))
                xLeft = t.Position[0];
            else if (up.Contains("GIỮA") || up.Contains("CEN") || up.Contains("MID"))
                xMid = t.Position[0];
            else if (up.Contains("PHẢI") || up.Contains("RIGHT"))
                xRight = t.Position[0];
        }

        // 3. Locate Row Headers and Contents
        var details = new BeamSectionDetailsDto
        {
            WidthMm = detectedBeams[0].b,
            HeightMm = detectedBeams[0].h
        };

        // Extract Level if present
        foreach (var t in rawEntities)
        {
            var mLev = LevelRegex().Match(t.Text);
            if (mLev.Success)
            {
                string sign = mLev.Groups["sign"].Value == "-" ? "-" : "";
                if (double.TryParse(sign + mLev.Groups["val"].Value, out double lev))
                {
                    details.TopElevM = lev;
                }
            }
        }

        // Parse rebar rows by finding text items
        var textEntities = rawEntities
            .Where(e => !string.IsNullOrWhiteSpace(e.Text) && e.Position.Length >= 2)
            .OrderByDescending(e => e.Position[1]) // top to bottom
            .ToList();

        // Group rows by approximate Y (within 50 units in CAD coords)
        var rowGroups = new List<List<CadEntityItemDto>>();
        foreach (var t in textEntities)
        {
            var grp = rowGroups.FirstOrDefault(g => Math.Abs(g[0].Position[1] - t.Position[1]) < 50.0);
            if (grp != null) grp.Add(t);
            else rowGroups.Add(new List<CadEntityItemDto> { t });
        }

        foreach (var grp in rowGroups)
        {
            // Sort items in this row from left to right (by X)
            grp.Sort((a, b) => a.Position[0].CompareTo(b.Position[0]));
            string rowHeader = grp[0].Text.ToUpperInvariant();

            if (rowHeader.Contains("THÉP TRÊN") || rowHeader.Contains("TOP BAR"))
            {
                var values = ExtractRowValuesByColumn(grp, xLeft, xMid, xRight);
                details.TopBarLeft = values.left;
                details.TopBarMid = values.mid;
                details.TopBarRight = values.right;
            }
            else if (rowHeader.Contains("THÉP DƯỚI") || rowHeader.Contains("BOTTOM BAR"))
            {
                var values = ExtractRowValuesByColumn(grp, xLeft, xMid, xRight);
                details.BotBarLeft = values.left;
                details.BotBarMid = values.mid;
                details.BotBarRight = values.right;
            }
            else if (rowHeader.Contains("THÉP ĐAI") || rowHeader.Contains("STIRRUP"))
            {
                var values = ExtractRowValuesByColumn(grp, xLeft, xMid, xRight);
                details.StirrupLeft = values.left;
                details.StirrupMid = values.mid;
                details.StirrupRight = values.right;
            }
            else if (rowHeader.Contains("THÉP GIÁ") || rowHeader.Contains("SUPPORT BAR"))
            {
                var values = ExtractRowValuesByColumn(grp, xLeft, xMid, xRight);
                details.SupportBarLeft = values.left;
                details.SupportBarMid = values.mid;
                details.SupportBarRight = values.right;
            }
            else if (rowHeader.Contains("THÉP BÊN") || rowHeader.Contains("SIDE BAR"))
            {
                var values = ExtractRowValuesByColumn(grp, xLeft, xMid, xRight);
                details.SideBar = !string.IsNullOrEmpty(values.left) ? values.left : (!string.IsNullOrEmpty(values.mid) ? values.mid : values.right);
            }
        }

        // 4. Construct Output for each detected beam
        foreach (var (bName, bWidth, bHeight, bElev) in detectedBeams)
        {
            var beamDetails = new BeamSectionDetailsDto
            {
                BeamName = bName,
                WidthMm = bWidth,
                HeightMm = bHeight,
                TopElevM = details.TopElevM,
                TopBarLeft = details.TopBarLeft,
                TopBarMid = details.TopBarMid,
                TopBarRight = details.TopBarRight,
                BotBarLeft = details.BotBarLeft,
                BotBarMid = details.BotBarMid,
                BotBarRight = details.BotBarRight,
                StirrupLeft = details.StirrupLeft,
                StirrupMid = details.StirrupMid,
                StirrupRight = details.StirrupRight,
                SideBar = details.SideBar,
                SupportBarLeft = details.SupportBarLeft,
                SupportBarMid = details.SupportBarMid,
                SupportBarRight = details.SupportBarRight
            };

            var el = new ElementDataDto
            {
                ElementId = $"{req.Building}-{req.Floor}-BEAM-{bName}",
                Building = req.Building,
                Group = "BEAM",
                Mark = bName,
                Description = $"Dầm {bName} {bWidth}x{bHeight}",
                LengthM = req.DefaultLengthM,
                WidthM = bWidth / 1000.0,
                HeightM = bHeight / 1000.0,
                TopElevM = details.TopElevM,
                Qty = 1,
                RebarTemplate = "BEAM_PARAMETRIC"
            };

            var rebarRows = BuildRebarConfigRows(el, beamDetails);

            resp.Beams.Add(new ParsedBeamResultDto
            {
                Element = el,
                RebarRows = rebarRows,
                SectionDetails = beamDetails
            });
        }

        resp.BeamCount = resp.Beams.Count;
        resp.Messages.Add($"Đã nhận diện thành công {resp.BeamCount} dầm ({string.Join(", ", detectedBeams.Select(b => b.name))}).");
        return resp;
    }

    private static (string left, string mid, string right) ExtractRowValuesByColumn(
        List<CadEntityItemDto> rowItems, double xLeft, double xMid, double xRight)
    {
        string left = "", mid = "", right = "";

        if (double.IsNaN(xLeft) || double.IsNaN(xMid) || double.IsNaN(xRight))
        {
            // Divide non-header items evenly across 3 slots
            var nonHeaders = rowItems.Skip(1).ToList();
            if (nonHeaders.Count >= 3)
            {
                left = nonHeaders[0].Text;
                mid = nonHeaders[1].Text;
                right = nonHeaders[2].Text;
            }
            else if (nonHeaders.Count == 1)
            {
                left = mid = right = nonHeaders[0].Text;
            }
            return (left, mid, right);
        }

        double split1 = (xLeft + xMid) / 2.0;
        double split2 = (xMid + xRight) / 2.0;

        foreach (var item in rowItems.Skip(1)) // Skip label column
        {
            double x = item.Position[0];
            if (x < split1)
            {
                left = string.IsNullOrEmpty(left) ? item.Text : $"{left}  {item.Text}";
            }
            else if (x < split2)
            {
                mid = string.IsNullOrEmpty(mid) ? item.Text : $"{mid}  {item.Text}";
            }
            else
            {
                right = string.IsNullOrEmpty(right) ? item.Text : $"{right}  {item.Text}";
            }
        }

        return (left, mid, right);
    }

    private static List<RebarConfigRowDto> BuildRebarConfigRows(ElementDataDto el, BeamSectionDetailsDto d)
    {
        var rows = new List<RebarConfigRowDto>();

        // 1. Top Rebar (Thép chủ trên)
        // Check Mid text for main continuous throughout bars
        string topMidText = !string.IsNullOrEmpty(d.TopBarMid) ? d.TopBarMid : d.TopBarLeft;
        var midMatches = BarCountAndDiaRegex().Matches(topMidText);
        double topQty = midMatches.Count > 0 ? double.Parse(midMatches[0].Groups["count"].Value) : 7;
        double topDia = midMatches.Count > 0 ? double.Parse(midMatches[0].Groups["dia"].Value) : 22;

        // Main top continuous throughout bar: Shape D07 (Thanh chữ U bẻ 2 mỏ vuông 35D)
        rows.Add(new RebarConfigRowDto
        {
            ElementId = el.ElementId,
            Building = el.Building,
            Group = el.Group,
            BarId = "T1",
            Role = "TOP",
            ShapeCode = "D07", // Chữ U bẻ 2 mỏ vuông
            Dia = topDia,
            BarDir = "X",
            DistDir = "FIXED",
            OverrideQtyExpr = topQty.ToString("0"),
            HookLExpr = "35",
            HookRExpr = "35",
            CoverExpr = "SPEC_Cover_Beam",
            Note = $"Thép chủ trên suốt ({topQty:0}D{topDia:0})"
        });

        // 1b. Top Reinforcement Bar at Left Support (Gia cường Gối Trái lớp 2)
        var leftMatches = BarCountAndDiaRegex().Matches(d.TopBarLeft);
        if (leftMatches.Count > 1 || (leftMatches.Count == 1 && double.Parse(leftMatches[0].Groups["count"].Value) > topQty))
        {
            double extraLeftQty = leftMatches.Count > 1
                ? double.Parse(leftMatches[1].Groups["count"].Value)
                : (double.Parse(leftMatches[0].Groups["count"].Value) - topQty);
            double extraLeftDia = leftMatches.Count > 1
                ? double.Parse(leftMatches[1].Groups["dia"].Value)
                : topDia;

            if (extraLeftQty > 0)
            {
                rows.Add(new RebarConfigRowDto
                {
                    ElementId = el.ElementId,
                    Building = el.Building,
                    Group = el.Group,
                    BarId = "T_GT",
                    Role = "TOP-Goi",
                    ShapeCode = "D06", // Chữ L mỏ vuông neo vào cột
                    Dia = extraLeftDia,
                    BarDir = "X*0.25",
                    DistDir = "FIXED",
                    OverrideQtyExpr = extraLeftQty.ToString("0"),
                    HookLExpr = "35",
                    HookRExpr = "0",
                    CoverExpr = "SPEC_Cover_Beam",
                    Note = $"Thép gia cường gối trái lớp 2 ({extraLeftQty:0}D{extraLeftDia:0})"
                });
            }
        }

        // 1c. Top Reinforcement Bar at Right Support (Gia cường Gối Phải lớp 2)
        var rightMatches = BarCountAndDiaRegex().Matches(d.TopBarRight);
        if (rightMatches.Count > 1 || (rightMatches.Count == 1 && double.Parse(rightMatches[0].Groups["count"].Value) > topQty))
        {
            double extraRightQty = rightMatches.Count > 1
                ? double.Parse(rightMatches[1].Groups["count"].Value)
                : (double.Parse(rightMatches[0].Groups["count"].Value) - topQty);
            double extraRightDia = rightMatches.Count > 1
                ? double.Parse(rightMatches[1].Groups["dia"].Value)
                : topDia;

            if (extraRightQty > 0)
            {
                rows.Add(new RebarConfigRowDto
                {
                    ElementId = el.ElementId,
                    Building = el.Building,
                    Group = el.Group,
                    BarId = "T_GP",
                    Role = "TOP-Goi",
                    ShapeCode = "D06", // Chữ L mỏ vuông neo vào cột
                    Dia = extraRightDia,
                    BarDir = "X*0.25",
                    DistDir = "FIXED",
                    OverrideQtyExpr = extraRightQty.ToString("0"),
                    HookLExpr = "0",
                    HookRExpr = "35",
                    CoverExpr = "SPEC_Cover_Beam",
                    Note = $"Thép gia cường gối phải lớp 2 ({extraRightQty:0}D{extraRightDia:0})"
                });
            }
        }

        // 2. Bottom Rebar (Thép chủ dưới)
        string botMidText = !string.IsNullOrEmpty(d.BotBarMid) ? d.BotBarMid : d.BotBarLeft;
        var botMatches = BarCountAndDiaRegex().Matches(botMidText);
        double botQty = botMatches.Count > 0 ? double.Parse(botMatches[0].Groups["count"].Value) : 7;
        double botDia = botMatches.Count > 0 ? double.Parse(botMatches[0].Groups["dia"].Value) : 20;

        rows.Add(new RebarConfigRowDto
        {
            ElementId = el.ElementId,
            Building = el.Building,
            Group = el.Group,
            BarId = "B1",
            Role = "BOT",
            ShapeCode = "D07", // Chữ U bẻ 2 mỏ vuông 35D
            Dia = botDia,
            BarDir = "X",
            DistDir = "FIXED",
            OverrideQtyExpr = botQty.ToString("0"),
            HookLExpr = "35",
            HookRExpr = "35",
            CoverExpr = "SPEC_Cover_Beam",
            Note = $"Thép chủ dưới suốt ({botQty:0}D{botDia:0})"
        });

        // 3. Side Bars (Thép bên hông chống phình/co ngót)
        if (!string.IsNullOrEmpty(d.SideBar))
        {
            var mSide = BarCountAndDiaRegex().Match(d.SideBar);
            double sideQty = mSide.Success ? double.Parse(mSide.Groups["count"].Value) : 2;
            double sideDia = mSide.Success ? double.Parse(mSide.Groups["dia"].Value) : 18;

            rows.Add(new RebarConfigRowDto
            {
                ElementId = el.ElementId,
                Building = el.Building,
                Group = el.Group,
                BarId = "S1",
                Role = "SIDE",
                ShapeCode = "D01", // Thanh thẳng
                Dia = sideDia,
                BarDir = "X",
                DistDir = "FIXED",
                OverrideQtyExpr = sideQty.ToString("0"),
                HookLExpr = "0",
                HookRExpr = "0",
                CoverExpr = "SPEC_Cover_Beam",
                Note = $"Thép bên chống phình/co ngót ({sideQty:0}D{sideDia:0})"
            });
        }

        // 4. Support Bar (Thép giá đỡ lớp 2)
        string suppText = !string.IsNullOrEmpty(d.SupportBarLeft) ? d.SupportBarLeft : (!string.IsNullOrEmpty(d.SupportBarMid) ? d.SupportBarMid : d.SupportBarRight);
        if (!string.IsNullOrEmpty(suppText))
        {
            var mSupp = SupportBarRegex().Match(suppText);
            if (mSupp.Success)
            {
                double gDia = double.Parse(mSupp.Groups["dia"].Value);
                double gSpacing = double.Parse(mSupp.Groups["spacing"].Value);
                rows.Add(new RebarConfigRowDto
                {
                    ElementId = el.ElementId,
                    Building = el.Building,
                    Group = el.Group,
                    BarId = "G1",
                    Role = "CHAIR",
                    ShapeCode = "D01",
                    Dia = gDia,
                    BarDir = "Y",
                    DistDir = "X",
                    Space1Mm = gSpacing,
                    CoverExpr = "SPEC_Cover_Beam",
                    Note = $"Thép giá đỡ lớp 2 (D{gDia:0}a{gSpacing:0})"
                });
            }
        }

        // 5. Stirrups (Thép đai ngoài A & đai giằng B, C, D)
        double stirDia = 12;
        double spaceGoi = 100;
        double spaceNhip = 150;

        var mStirGoi = StirrupSpecRegex().Match(d.StirrupLeft);
        if (mStirGoi.Success)
        {
            if (double.TryParse(mStirGoi.Groups["dia"].Value, out double sd)) stirDia = sd;
            if (double.TryParse(mStirGoi.Groups["spacing"].Value, out double sg)) spaceGoi = sg;
        }
        var mStirNhip = StirrupSpecRegex().Match(d.StirrupMid);
        if (mStirNhip.Success)
        {
            if (double.TryParse(mStirNhip.Groups["spacing"].Value, out double sn)) spaceNhip = sn;
            else spaceNhip = spaceGoi;
        }
        else spaceNhip = spaceGoi;

        // 5a. Outer Stirrup A (Đai kín chữ nhật chu vi D39)
        rows.Add(new RebarConfigRowDto
        {
            ElementId = el.ElementId,
            Building = el.Building,
            Group = el.Group,
            BarId = "DA",
            Role = "STIR.R",
            ShapeCode = "D39", // Thép đai kín chữ nhật chu vi chuẩn
            Dia = stirDia,
            BarDir = "PERIM_YH",
            DistDir = "ZONE2_X",
            Space1Mm = spaceGoi,
            Space2Mm = spaceNhip,
            CoverExpr = "SPEC_Cover_Beam",
            Note = $"Đai ngoài kín A (D{stirDia:0}a{spaceGoi:0}/{spaceNhip:0})"
        });

        // 5b. Inner Horizontal Ties B, C (Đai giằng ngang D24)
        string stirAll = $"{d.StirrupLeft} {d.StirrupMid} {d.StirrupRight}".ToUpperInvariant();
        bool hasHorizontalTie = stirAll.Contains("B") || stirAll.Contains("C") || el.HeightM >= 0.70;
        if (hasHorizontalTie)
        {
            int hTieCount = el.HeightM >= 0.90 ? 2 : 1;
            rows.Add(new RebarConfigRowDto
            {
                ElementId = el.ElementId,
                Building = el.Building,
                Group = el.Group,
                BarId = "DBC",
                Role = "STIR.C",
                ShapeCode = "D24", // Đai C móc 135° 2 đầu
                Dia = stirDia,
                LayerMult = hTieCount,
                BarDir = "TIE_Y",
                DistDir = "ZONE2_X",
                Space1Mm = spaceGoi,
                Space2Mm = spaceNhip,
                CoverExpr = "SPEC_Cover_Beam",
                Note = $"{hTieCount} Đai giằng ngang (B, C: D{stirDia:0}a{spaceGoi:0}/{spaceNhip:0})"
            });
        }

        // 5c. Inner Vertical Ties D (Đai giằng đứng D24 theo chiều cao dầm)
        bool hasVerticalTie = stirAll.Contains("D") || el.WidthM >= 0.80;
        if (hasVerticalTie)
        {
            // Beam 1300mm wide -> 4 vertical tie branches (Image 0)
            int vTieBranches = el.WidthM >= 1.20 ? 4 : (el.WidthM >= 0.90 ? 3 : 2);
            rows.Add(new RebarConfigRowDto
            {
                ElementId = el.ElementId,
                Building = el.Building,
                Group = el.Group,
                BarId = "DD",
                Role = "STIR.C",
                ShapeCode = "D24", // Đai C móc 135° 2 đầu dọc chiều đứng
                Dia = stirDia,
                LayerMult = vTieBranches,
                BarDir = "H", // Dọc theo chiều cao H
                DistDir = "ZONE2_X",
                Space1Mm = spaceGoi,
                Space2Mm = spaceNhip,
                CoverExpr = "SPEC_Cover_Beam",
                Note = $"{vTieBranches} Đai giằng đứng (D: D{stirDia:0}a{spaceGoi:0}/{spaceNhip:0})"
            });
        }

        return rows;
    }
}
