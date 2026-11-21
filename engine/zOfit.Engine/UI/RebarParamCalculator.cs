using System.Globalization;
using ZOfit.Engine.Models;

namespace ZOfit.Engine.UI;

public sealed record RebarCalcPreview(
    double LeftMm,
    double Le1aMm,
    double Le1bMm,
    double MidMm,
    double Ri2bMm,
    double Ri2aMm,
    double RightMm,
    double QtyBar,
    double LbarMm,
    double LapMm,
    double SumLenM,
    double DensityKgM,
    double WeightKg
);

public static class RebarParamCalculator
{
    public static double ResolveCoverMm(ProjectSpecDto spec, ElementDataDto el, string? coverExpr)
    {
        if (!string.IsNullOrWhiteSpace(coverExpr))
        {
            string s = coverExpr.Trim().TrimStart('=').ToUpperInvariant();
            if (double.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out double num) && num >= 0)
                return num;
            if (s.Contains("SPEC_COVER_FTG")) return spec.CoverFtgMm;
            if (s.Contains("SPEC_COVER_COL")) return spec.CoverColMm;
            if (s.Contains("SPEC_COVER_BEAM")) return spec.CoverBeamMm;
            if (s.Contains("SPEC_COVER_WALL")) return spec.CoverWallMm;
            if (s.Contains("SPEC_COVER_SLAB")) return spec.CoverSlabMm;
        }

        return (el.Group ?? "").Trim().ToLowerInvariant() switch
        {
            "foundation" or "groundbeam" or "slabongrade" => spec.CoverFtgMm,
            "column" => spec.CoverColMm,
            "beam" or "lintel" => spec.CoverBeamMm,
            "wall" or "tank" or "manhole" => spec.CoverWallMm,
            _ => spec.CoverSlabMm
        };
    }

    public static double ResolveHookMm(string? expr, double diaMm)
    {
        if (string.IsNullOrWhiteSpace(expr)) return 0;
        if (double.TryParse(expr.Trim(), NumberStyles.Any, CultureInfo.InvariantCulture, out double v))
        {
            if (v <= 0) return 0;
            return v < 50 ? Math.Round(diaMm * v, 0) : Math.Round(v, 0);
        }
        return 0;
    }

    public static double EvalOptionalExpr(string? expr, ProjectSpecDto spec, double diaMm)
    {
        if (string.IsNullOrWhiteSpace(expr)) return 0;
        string s = expr.Trim().TrimStart('=').ToUpperInvariant();
        if (double.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out double direct))
            return direct;

        // Evaluate common parametric spec expressions stored in ±Lth / L.m
        if (s.Contains("SPEC_LLAP_MULT") && s.Contains("SPEC_COL_STAGGEROFFSET"))
            return spec.LlapMult * diaMm + spec.ColStaggerOffsetMm;
        if (s.Contains("SPEC_LLAP_MULT"))
            return spec.LlapMult * diaMm;
        if (s.Contains("SPEC_LAN_TENSION") && s.Contains("2*MAX"))
            return 2 * Math.Max(250, spec.LanTensionMult * diaMm);
        if (s.Contains("SPEC_LAN_TENSION"))
            return spec.LanTensionMult * diaMm;
        if (s.Contains("SPEC_LAN_COMP"))
            return spec.LanCompMult * diaMm;
        if (s.Contains("SPEC_PILE_CUTHEAD"))
            return spec.PileCutHeadM * 1000 + 1000;
        if (s.Contains("PI()"))
            return Math.Round(2 * Math.PI * 150 * 8, 0);

        return 0;
    }

    public static RebarCalcPreview ComputeRow(ProjectSpecDto spec, ElementDataDto el, RebarConfigRowDto row)
    {
        double dia = row.Dia > 0 ? row.Dia : 14;
        double part = row.LayerMult > 0 ? row.LayerMult : 1;
        double eleQty = el.Qty > 0 ? el.Qty : 1;
        double x = el.LengthM * 1000.0;
        double y = el.WidthM * 1000.0;
        double h = el.HeightM * 1000.0;
        double bv = ResolveCoverMm(spec, el, row.CoverExpr);

        double left = ResolveHookMm(row.HookLExpr, dia);
        double le1a = row.Le1aMm ?? 0;
        double le1b = row.Le1bMm ?? 0;
        double ri2b = row.Ri2bMm ?? 0;
        double ri2a = row.Ri2aMm ?? 0;
        double right = ResolveHookMm(row.HookRExpr, dia);

        double mid;
        if (!string.IsNullOrWhiteSpace(row.OverrideLenExpr))
        {
            mid = EvalOptionalExpr(row.OverrideLenExpr, spec, dia);
        }
        else
        {
            double baseMid = (row.BarDir ?? "X").Trim().ToUpperInvariant() switch
            {
                "X" => x - 2 * bv,
                "Y" => y - 2 * bv,
                "H" => h - 2 * bv,
                "X*0.25" => Math.Round(x * 0.25, 0, MidpointRounding.AwayFromZero),
                "Y*0.25" => Math.Round(y * 0.25, 0, MidpointRounding.AwayFromZero),
                "X*0.7" => Math.Round(x * 0.7, 0, MidpointRounding.AwayFromZero),
                "PERIM_YH" => 2 * (y + h - 4 * bv),
                "PERIM_XY" => 2 * (x + y - 4 * bv),
                "PERIM_XY_INNER" => 2 * (Math.Round(x * 0.67, 0, MidpointRounding.AwayFromZero) + y - 4 * bv),
                "TIE_Y" => y - 2 * bv,
                "TIE_X" => x - 2 * bv,
                "SLOPED_XH" => Math.Round(Math.Sqrt(x * x + h * h), 0, MidpointRounding.AwayFromZero),
                "CHAIR_H" => Math.Max(120, (h - 2 * bv) * 2 + 150),
                _ => 0
            };
            double dLth = EvalOptionalExpr(row.DeltaLenExpr, spec, dia);
            mid = Math.Max(0, baseMid + dLth);
        }

        double qtyBar;
        if (!string.IsNullOrWhiteSpace(row.OverrideQtyExpr) &&
            double.TryParse(row.OverrideQtyExpr.Trim(), NumberStyles.Any, CultureInfo.InvariantCulture, out double qOv))
        {
            qtyBar = qOv;
        }
        else
        {
            double dp = EvalOptionalExpr(row.DeltaDistExpr, spec, dia);
            double s1 = Math.Max(1, row.Space1Mm ?? 200);
            double s2 = Math.Max(1, row.Space2Mm ?? s1);
            qtyBar = (row.DistDir ?? "FIXED").Trim().ToUpperInvariant() switch
            {
                "X" => Math.Round((x + dp) / s1, 0, MidpointRounding.AwayFromZero) + 1,
                "Y" => Math.Round((y + dp) / s1, 0, MidpointRounding.AwayFromZero) + 1,
                "H" => Math.Round((h + dp) / s1, 0, MidpointRounding.AwayFromZero) + 1,
                "ZONE2_X" => Math.Round(((x + dp) * 0.5) / s1, 0, MidpointRounding.AwayFromZero)
                           + Math.Round(((x + dp) * 0.5) / s2, 0, MidpointRounding.AwayFromZero) + 2,
                "ZONE2_H" => Math.Round(((h + dp) * 0.5) / s1, 0, MidpointRounding.AwayFromZero)
                           + Math.Round(((h + dp) * 0.5) / s2, 0, MidpointRounding.AwayFromZero) + 1,
                "XY_GRID" => (Math.Round((x + dp) / s1, 0, MidpointRounding.AwayFromZero) + 1)
                           * (Math.Round((y + dp) / s2, 0, MidpointRounding.AwayFromZero) + 1),
                "XH_GRID" => (Math.Round((x + dp) / s1, 0, MidpointRounding.AwayFromZero) + 1)
                           * (Math.Round((h + dp) / s2, 0, MidpointRounding.AwayFromZero) + 1),
                _ => 1
            };
        }

        double lbar = left + le1a + le1b + mid + ri2b + ri2a + right;
        double lap = 0;
        if (dia >= 10 && dia < spec.CouplerMinD && spec.StdBarLenMm > 0)
        {
            double laps = Math.Max(0, Math.Ceiling(lbar / spec.StdBarLenMm) - 1);
            lap = laps * dia * spec.LlapMult;
        }
        double sumLenM = Math.Round(eleQty * part * qtyBar * (lbar + lap) / 1000.0, 2, MidpointRounding.AwayFromZero);
        double density = Math.Round(Math.PI * Math.Pow(dia / 1000.0, 2) / 4.0 * 7850.0, 3, MidpointRounding.AwayFromZero);
        double kg = Math.Round(sumLenM * density, 2, MidpointRounding.AwayFromZero);

        return new RebarCalcPreview(left, le1a, le1b, mid, ri2b, ri2a, right, qtyBar, lbar, lap, sumLenM, density, kg);
    }

    public static RebarConfigRowDto CreateQuickRoleRow(ElementDataDto el, ProjectSpecDto spec, string roleKey, int seqNo)
    {
        string mk = string.IsNullOrWhiteSpace(el.Mark) ? el.ElementId : el.Mark.Trim();
        string bvExpr = "=" + ResolveCoverName(el.Group);
        RebarConfigRowDto Make(string suffix, string role, string shape, double dia, double part, string barDir, string distDir,
            string hookL, string hookR, double? sp1, double? sp2, string ovQty, string dLth, string note) => new()
        {
            RowKind = "BAR",
            ElementId = el.ElementId,
            Building = el.Building,
            Group = el.Group,
            BarId = $"{mk}-{suffix}{seqNo}",
            Role = role,
            ShapeCode = shape,
            Dia = dia,
            LayerMult = part,
            BarDir = barDir,
            DistDir = distDir,
            HookLExpr = hookL,
            CoverExpr = bvExpr,
            HookRExpr = hookR,
            Space1Mm = sp1,
            Space2Mm = sp2,
            OverrideQtyExpr = ovQty,
            DeltaLenExpr = dLth,
            Note = note
        };

        return roleKey.ToUpperInvariant() switch
        {
            "BOT.X" => Make("BX", "BOT.X", "D07", 14, 1, "X", "Y", "150", "150", 150, null, "", "", "Thep lop duoi phuong X"),
            "BOT.Y" => Make("BY", "BOT.Y", "D07", 14, 1, "Y", "X", "150", "150", 150, null, "", "", "Thep lop duoi phuong Y"),
            "TOP.X" => Make("TX", "TOP.X", "D07", 12, 1, "X", "Y", "120", "120", 200, null, "", "", "Thep lop tren phuong X"),
            "TOP.Y" => Make("TY", "TOP.Y", "D07", 12, 1, "Y", "X", "120", "120", 200, null, "", "", "Thep lop tren phuong Y"),
            "AROUND" => Make("AR", "AROUND", "D02", 12, 1, "X", "H", "150", "150", 300, null, "", "", "Thep dai quay hong mong"),
            "TOP" => Make("T", "TOP", "D07", 18, 1, "X", "FIXED", "20", "20", null, null, "4", "", "Thep chu mien tren (mo 20D)"),
            "BOT" => Make("B", "BOT", "D07", 18, 1, "X", "FIXED", "20", "20", null, null, "4", "", "Thep chu mien duoi (mo 20D)"),
            "SIDE" => Make("SK", "SIDE", "D01", 12, 2, "X", "FIXED", "0", "0", null, null, "1", "=2*MAX(250,SPEC_Lan_Tension*L{r})", "Thep gia thanh dam (2 mat)"),
            "TOP-GOI" => Make("GG", "TOP-Goi", "D06", 18, 2, "X*0.25", "FIXED", "20", "0", null, null, "2", "=SPEC_Lan_Tension*L{r}", "Gia cuong goi L/4 (2 dau)"),
            "BOT-NHIP" => Make("GN", "BOT-Nhip", "D01", 18, 1, "X*0.7", "FIXED", "0", "0", null, null, "2", "", "Gia cuong bung nhip 0.7L"),
            "STIR.BEAM" => Make("SB", "STIR.R", "D39", 10, 1, "PERIM_YH", "ZONE2_X", "10", "10", 150, 200, "", "", "Dai kin dam (@1 goi / @2 giua)"),
            "STIR.COL" => Make("SC", "STIR.R", "D39", 10, 1, "PERIM_XY", "ZONE2_H", "10", "10", 100, 200, "", "", "Dai bao ngoai cot (@1 noi / @2 giua)"),
            "STIR.IN" => Make("SI", "STIR.IN", "D39", 10, 1, "PERIM_XY_INNER", "ZONE2_H", "10", "10", 100, 200, "", "", "Dai long trong cot"),
            "TIE.C.BEAM" => Make("TC", "STIR.C", "D24", 8, 1, "TIE_Y", "ZONE2_X", "10", "10", 150, 200, "", "", "Dai C moc giua dam"),
            "TIE.C.COL" => Make("TC", "STIR.C", "D24", 8, 1, "TIE_Y", "ZONE2_H", "10", "10", 100, 200, "", "", "Dai C cot"),
            "MAIN.A" => Make("CA", "MAIN.A", "D06", 18, 1, "H", "FIXED", "30", "0", null, null, "4", "=SPEC_Llap_Mult*L{r}+SPEC_Col_StaggerOffset", "Thep chu nhom A (noi sole +Stagger)"),
            "MAIN.B" => Make("CB", "MAIN.B", "D06", 18, 1, "H", "FIXED", "30", "0", null, null, "4", "=SPEC_Llap_Mult*L{r}", "Thep chu nhom B (noi chuan)"),
            "MAIN.V" => Make("VD", "MAIN.V", "D06", 14, 2, "H", "X", "40", "0", 200, null, "", "", "Thep dung vach 2 mat (rai doc L)"),
            "MAIN.H" => Make("HN", "MAIN.H", "D07", 12, 2, "X", "H", "100", "100", 200, null, "", "", "Thep ngang vach 2 mat (rai doc H)"),
            "TIE.C.WALL" => Make("DW", "TIE.C", "D24", 8, 1, "TIE_Y", "XH_GRID", "10", "10", 600, 600, "", "", "Dai C giu 2 mat vach (@600x@600)"),
            "HAT.X" => Make("HX", "HAT.X", "D07", 12, 2, "X*0.25", "Y", "100", "100", 150, null, "", "=SPEC_Lan_Tension*L{r}", "Thep mu goi san phuong X"),
            "HAT.Y" => Make("HY", "HAT.Y", "D07", 12, 2, "Y*0.25", "X", "100", "100", 150, null, "", "=SPEC_Lan_Tension*L{r}", "Thep mu goi san phuong Y"),
            "CHAIR" => Make("CK", "CHAIR", "D07", 10, 1, "CHAIR_H", "XY_GRID", "12", "12", 600, 600, "", "", "Chan cho ke thep lop tren (@600x@600)"),
            "ST-01" => Make("S1", "ST-01", "D15", 12, 1, "SLOPED_XH", "Y", "250", "450", 150, null, "", "", "Thep chu ban thang nghieng"),
            "ST-03" => Make("S3", "ST-03", "D06", 10, 1, "X*0.25", "Y", "120", "0", 150, null, "", "", "Thep mu goi ban thang"),
            "ST-04" => Make("S4", "ST-04", "D01", 8, 1, "Y", "X", "0", "0", 200, null, "", "", "Thep phan bo ngang ban thang"),
            _ => Make("SP", "CUSTOM", "D07", 14, 1, "X", "FIXED", "150", "150", null, null, "2", "", "Thanh thep tuy bien")
        };
    }

    private static string ResolveCoverName(string? group) => (group ?? "").Trim().ToLowerInvariant() switch
    {
        "foundation" or "groundbeam" or "slabongrade" => "SPEC_Cover_Ftg",
        "column" => "SPEC_Cover_Col",
        "beam" or "lintel" => "SPEC_Cover_Beam",
        "wall" or "tank" or "manhole" => "SPEC_Cover_Wall",
        _ => "SPEC_Cover_Slab"
    };
}
