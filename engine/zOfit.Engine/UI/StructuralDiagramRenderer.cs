using System.Drawing;
using System.Drawing.Drawing2D;
using ZOfit.Engine.Models;

namespace ZOfit.Engine.UI;

public static class StructuralDiagramRenderer
{
    private static readonly Color CanvasBg = Color.FromArgb(248, 251, 255);
    private static readonly Color GridColor = Color.FromArgb(232, 240, 248);
    private static readonly Color ConcFill = Color.FromArgb(218, 232, 224);
    private static readonly Color ConcStroke = Color.FromArgb(72, 102, 128);
    private static readonly Color LeanFill = Color.FromArgb(228, 228, 228);
    private static readonly Color DimColor = Color.FromArgb(80, 100, 120);
    private static readonly Color BarNormal = Color.FromArgb(14, 92, 166);
    private static readonly Color BarSecondary = Color.FromArgb(24, 128, 92);
    private static readonly Color BarActive = Color.FromArgb(220, 42, 24);

    public static void RenderElementDiagram(
        Graphics g,
        Rectangle bounds,
        ElementDataDto el,
        IReadOnlyList<RebarConfigRowDto> rows,
        IReadOnlyList<RebarCalcPreview> calcs,
        int selectedIndex)
    {
        g.SmoothingMode = SmoothingMode.AntiAlias;
        g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.ClearTypeGridFit;
        g.Clear(CanvasBg);

        // Subtle drafting grid
        using (var gridPen = new Pen(GridColor, 1f))
        {
            for (int x = bounds.Left; x < bounds.Right; x += 24) g.DrawLine(gridPen, x, bounds.Top, x, bounds.Bottom);
            for (int y = bounds.Top; y < bounds.Bottom; y += 24) g.DrawLine(gridPen, bounds.Left, y, bounds.Right, y);
        }

        string grp = (el.Group ?? "").Trim().ToLowerInvariant();
        if (grp is "beam" or "groundbeam" or "lintel")
            DrawBeamDiagram(g, bounds, el, rows, calcs, selectedIndex);
        else if (grp is "column")
            DrawColumnDiagram(g, bounds, el, rows, calcs, selectedIndex);
        else if (grp is "wall")
            DrawWallDiagram(g, bounds, el, rows, calcs, selectedIndex);
        else if (grp is "slab" or "slabongrade")
            DrawSlabDiagram(g, bounds, el, rows, calcs, selectedIndex);
        else if (grp is "stair")
            DrawStairDiagram(g, bounds, el, rows, calcs, selectedIndex);
        else
            DrawFootingDiagram(g, bounds, el, rows, calcs, selectedIndex);
    }

    private static bool IsMatch(RebarConfigRowDto r, params string[] keys)
    {
        string role = (r.Role ?? "").Trim().ToUpperInvariant();
        string bid = (r.BarId ?? "").Trim().ToUpperInvariant();
        return keys.Any(k => role.Contains(k, StringComparison.OrdinalIgnoreCase) || bid.Contains(k, StringComparison.OrdinalIgnoreCase));
    }

    private static (bool present, bool active, string tag) FindRoleState(
        IReadOnlyList<RebarConfigRowDto> rows, IReadOnlyList<RebarCalcPreview> calcs, int selIdx, params string[] keys)
    {
        for (int i = 0; i < rows.Count; i++)
        {
            if (IsMatch(rows[i], keys))
            {
                bool act = i == selIdx;
                string info = $"Ø{(int)rows[i].Dia} ({calcs[i].QtyBar:0}c, L={calcs[i].LbarMm:0})";
                return (true, act, $"[{i + 1}] {rows[i].Role}: {info}");
            }
        }
        return (false, false, "");
    }

    private static Pen BarPen(bool active, Color baseCol, float w = 2.2f)
        => new(active ? BarActive : baseCol, active ? 3.5f : w) { LineJoin = LineJoin.Round };

    private static void DrawDimLine(Graphics g, float x1, float y1, float x2, float y2, string label)
    {
        using var p = new Pen(DimColor, 1f);
        using var f = new Font("Segoe UI", 8f, FontStyle.Bold);
        using var br = new SolidBrush(DimColor);
        g.DrawLine(p, x1, y1, x2, y2);
        // ticks
        bool horiz = Math.Abs(y1 - y2) < 1f;
        if (horiz)
        {
            g.DrawLine(p, x1, y1 - 4, x1, y1 + 4);
            g.DrawLine(p, x2, y2 - 4, x2, y2 + 4);
            var sz = g.MeasureString(label, f);
            g.FillRectangle(new SolidBrush(CanvasBg), (x1 + x2 - sz.Width) / 2, y1 - sz.Height - 1, sz.Width + 2, sz.Height);
            g.DrawString(label, f, br, (x1 + x2 - sz.Width) / 2, y1 - sz.Height - 1);
        }
        else
        {
            g.DrawLine(p, x1 - 4, y1, x1 + 4, y1);
            g.DrawLine(p, x2 - 4, y2, x2 + 4, y2);
            var sz = g.MeasureString(label, f);
            g.DrawString(label, f, br, x1 + 4, (y1 + y2 - sz.Height) / 2);
        }
    }

    private static void DrawCallout(Graphics g, float tx, float ty, float lx, float ly, string text, bool active, bool alignRight = false)
    {
        if (string.IsNullOrWhiteSpace(text)) return;
        Color col = active ? BarActive : Color.FromArgb(22, 54, 92);
        using var pen = new Pen(col, active ? 1.6f : 1f);
        using var f = new Font("Segoe UI", 8f, active ? FontStyle.Bold : FontStyle.Regular);
        g.DrawLine(pen, tx, ty, lx, ly);
        g.FillEllipse(new SolidBrush(col), tx - 3, ty - 3, 6, 6);

        var sz = g.MeasureString(text, f);
        float rx = alignRight ? lx - sz.Width - 6 : lx + 2;
        float ry = ly - sz.Height / 2;
        var pill = new RectangleF(rx, ry, sz.Width + 6, sz.Height + 2);
        g.FillRectangle(new SolidBrush(active ? Color.FromArgb(255, 240, 235) : Color.FromArgb(240, 246, 255)), pill);
        g.DrawRectangle(pen, pill.X, pill.Y, pill.Width, pill.Height);
        g.DrawString(text, f, new SolidBrush(col), rx + 3, ry + 1);
    }

    private static void DrawFootingDiagram(Graphics g, Rectangle b, ElementDataDto el, IReadOnlyList<RebarConfigRowDto> rows, IReadOnlyList<RebarCalcPreview> calcs, int selIdx)
    {
        // Left: Plan View | Right: Cross Section
        float wPlan = Math.Min(195, b.Width * 0.36f);
        float hPlan = Math.Min(155, b.Height * 0.54f);
        var planRect = new RectangleF(b.Left + 44, b.Top + 46, wPlan, hPlan);
        var secRect = new RectangleF(planRect.Right + 65, b.Top + 68, Math.Min(210, b.Width * 0.36f), 115);

        using var titleF = new Font("Segoe UI", 8.5f, FontStyle.Bold);
        using var titleBr = new SolidBrush(ConcStroke);
        g.DrawString("MAT BANG MONG", titleF, titleBr, planRect.Left, b.Top + 12);
        g.DrawString("MAT CAT DUNG MONG", titleF, titleBr, secRect.Left, b.Top + 12);

        // Lean + Concrete Plan
        g.FillRectangle(new SolidBrush(LeanFill), planRect.Left - 8, planRect.Top - 8, planRect.Width + 16, planRect.Height + 16);
        g.FillRectangle(new SolidBrush(ConcFill), planRect);
        g.DrawRectangle(new Pen(ConcStroke, 1.5f), planRect.X, planRect.Y, planRect.Width, planRect.Height);
        // Pedestal stub
        var ped = new RectangleF(planRect.Left + planRect.Width * 0.38f, planRect.Top + planRect.Height * 0.38f, planRect.Width * 0.24f, planRect.Height * 0.24f);
        g.FillRectangle(Brushes.WhiteSmoke, ped);
        g.DrawRectangle(new Pen(ConcStroke, 1.2f), ped.X, ped.Y, ped.Width, ped.Height);

        var stBotX = FindRoleState(rows, calcs, selIdx, "BOT.X", "FTG_BOT_X", "M1", "BX");
        var stBotY = FindRoleState(rows, calcs, selIdx, "BOT.Y", "FTG_BOT_Y", "M2", "BY");
        var stTop = FindRoleState(rows, calcs, selIdx, "TOP", "TX", "TY");
        var stChair = FindRoleState(rows, calcs, selIdx, "CHAIR", "CK");
        var stAround = FindRoleState(rows, calcs, selIdx, "AROUND", "AR");

        // Rebar mesh in Plan
        using (var pX = BarPen(stBotX.active, BarNormal))
            for (int i = 1; i <= 4; i++)
            {
                float y = planRect.Top + i * planRect.Height / 5f;
                g.DrawLine(pX, planRect.Left + 8, y, planRect.Right - 8, y);
            }
        using (var pY = BarPen(stBotY.active, BarSecondary))
            for (int i = 1; i <= 4; i++)
            {
                float x = planRect.Left + i * planRect.Width / 5f;
                g.DrawLine(pY, x, planRect.Top + 8, x, planRect.Bottom - 8);
            }

        DrawDimLine(g, planRect.Left, planRect.Bottom + 20, planRect.Right, planRect.Bottom + 20, $"X = {el.LengthM * 1000:0} mm");
        DrawDimLine(g, planRect.Left - 18, planRect.Top, planRect.Left - 18, planRect.Bottom, $"Y={el.WidthM * 1000:0}");

        // Cross Section
        var leanSec = new RectangleF(secRect.Left - 10, secRect.Bottom, secRect.Width + 20, 14);
        g.FillRectangle(new SolidBrush(LeanFill), leanSec);
        g.DrawRectangle(new Pen(Color.Gray, 1f), leanSec.X, leanSec.Y, leanSec.Width, leanSec.Height);
        g.FillRectangle(new SolidBrush(ConcFill), secRect);
        g.DrawRectangle(new Pen(ConcStroke, 1.5f), secRect.X, secRect.Y, secRect.Width, secRect.Height);
        g.DrawRectangle(new Pen(ConcStroke, 1.2f), secRect.Left + secRect.Width * 0.38f, secRect.Top - 26, secRect.Width * 0.24f, 26);

        // Bot U-bar in Section
        using (var pB = BarPen(stBotX.active || stBotY.active, BarNormal, 2.5f))
        {
            float yBot = secRect.Bottom - 10;
            g.DrawLines(pB, new PointF[] {
                new(secRect.Left + 8, yBot - 28),
                new(secRect.Left + 8, yBot),
                new(secRect.Right - 8, yBot),
                new(secRect.Right - 8, yBot - 28)
            });
        }
        if (stTop.present)
        {
            using var pT = BarPen(stTop.active, BarSecondary, 2.2f);
            float yTop = secRect.Top + 10;
            g.DrawLines(pT, new PointF[] {
                new(secRect.Left + 8, yTop + 26),
                new(secRect.Left + 8, yTop),
                new(secRect.Right - 8, yTop),
                new(secRect.Right - 8, yTop + 26)
            });
        }
        if (stChair.present)
        {
            using var pC = BarPen(stChair.active, Color.DarkGoldenrod, 2f);
            float cx = secRect.Left + 34, yB = secRect.Bottom - 12, yT = secRect.Top + 12;
            g.DrawLines(pC, new PointF[] { new(cx - 10, yB), new(cx, yB), new(cx, yT), new(cx + 18, yT), new(cx + 18, yB), new(cx + 28, yB) });
        }
        if (stAround.present)
        {
            using var pA = BarPen(stAround.active, Color.Purple, 2f);
            float yMid = (secRect.Top + secRect.Bottom) / 2f;
            g.DrawLine(pA, secRect.Left + 6, yMid, secRect.Right - 6, yMid);
        }

        DrawDimLine(g, secRect.Right + 10, secRect.Top, secRect.Right + 10, secRect.Bottom, $"H={el.HeightM * 1000:0}");

        // Callouts
        if (stBotX.present) DrawCallout(g, planRect.Right - 20, planRect.Top + planRect.Height * 0.2f, secRect.Left - 5, b.Top + 38, stBotX.tag, stBotX.active, true);
        if (stBotY.present) DrawCallout(g, planRect.Left + planRect.Width * 0.4f, planRect.Bottom - 20, planRect.Right + 10, b.Bottom - 28, stBotY.tag, stBotY.active);
        if (stTop.present) DrawCallout(g, secRect.Left + secRect.Width * 0.7f, secRect.Top + 10, secRect.Left, b.Top + 44, stTop.tag, stTop.active);
        if (stAround.present) DrawCallout(g, secRect.Left + 18, (secRect.Top + secRect.Bottom) / 2, secRect.Left - 12, b.Bottom - 54, stAround.tag, stAround.active, true);
        if (stChair.present) DrawCallout(g, secRect.Left + 44, secRect.Top + 14, secRect.Right - 20, b.Bottom - 22, stChair.tag, stChair.active);
    }

    private static void DrawBeamDiagram(Graphics g, Rectangle b, ElementDataDto el, IReadOnlyList<RebarConfigRowDto> rows, IReadOnlyList<RebarCalcPreview> calcs, int selIdx)
    {
        using var tf = new Font("Segoe UI", 8.5f, FontStyle.Bold);
        using var sf = new Font("Segoe UI", 7.5f, FontStyle.Regular);
        using var tbr = new SolidBrush(ConcStroke);

        // Roles lookup
        var stTop = FindRoleState(rows, calcs, selIdx, "TOP", "BEAM_TOP");
        var stBot = FindRoleState(rows, calcs, selIdx, "BOT", "BEAM_BOT");
        var stSide = FindRoleState(rows, calcs, selIdx, "SIDE", "BEAM_SKIN");
        var stGoi = FindRoleState(rows, calcs, selIdx, "TOP-GOI", "SUP_ADD", "GG");
        var stNhip = FindRoleState(rows, calcs, selIdx, "BOT-NHIP", "MID_ADD", "GN");
        var stStir = FindRoleState(rows, calcs, selIdx, "STIR.R", "STIRRUP_RECT", "ST1");
        var stTieC = FindRoleState(rows, calcs, selIdx, "STIR.C", "STIRRUP_C", "ST2");

        // Top half: Longitudinal Elevation (Mặt cắt dọc)
        float elevW = Math.Min(340, b.Width * 0.55f);
        var elev = new RectangleF(b.Left + 28, b.Top + 42, elevW, 85);
        g.DrawString("MẶT CẮT DỌC DẦM / GIẰNG MÓNG", tf, tbr, elev.Left, b.Top + 10);

        g.FillRectangle(new SolidBrush(ConcFill), elev);
        g.DrawRectangle(new Pen(ConcStroke, 1.5f), elev.X, elev.Y, elev.Width, elev.Height);

        // Longitudinal Top / Bot
        if (stTop.present)
        {
            using var p = BarPen(stTop.active, BarNormal, 2.5f);
            g.DrawLines(p, new PointF[] { new(elev.Left + 8, elev.Top + 24), new(elev.Left + 8, elev.Top + 10), new(elev.Right - 8, elev.Top + 10), new(elev.Right - 8, elev.Top + 24) });
        }
        if (stBot.present)
        {
            using var p = BarPen(stBot.active, BarNormal, 2.5f);
            g.DrawLines(p, new PointF[] { new(elev.Left + 8, elev.Bottom - 24), new(elev.Left + 8, elev.Bottom - 10), new(elev.Right - 8, elev.Bottom - 10), new(elev.Right - 8, elev.Bottom - 24) });
        }
        if (stGoi.present)
        {
            using var p = BarPen(stGoi.active, BarSecondary, 2.4f);
            g.DrawLines(p, new PointF[] { new(elev.Left + 10, elev.Top + 28), new(elev.Left + 10, elev.Top + 16), new(elev.Left + elev.Width * 0.28f, elev.Top + 16) });
            g.DrawLines(p, new PointF[] { new(elev.Right - 10, elev.Top + 28), new(elev.Right - 10, elev.Top + 16), new(elev.Right - elev.Width * 0.28f, elev.Top + 16) });
        }
        if (stNhip.present)
        {
            using var p = BarPen(stNhip.active, Color.DarkMagenta, 2.4f);
            g.DrawLine(p, elev.Left + elev.Width * 0.16f, elev.Bottom - 16, elev.Right - elev.Width * 0.16f, elev.Bottom - 16);
        }
        if (stSide.present)
        {
            using var p = BarPen(stSide.active, Color.Teal, 1.8f);
            p.DashStyle = DashStyle.Dash;
            g.DrawLine(p, elev.Left + 10, (elev.Top + elev.Bottom) / 2f, elev.Right - 10, (elev.Top + elev.Bottom) / 2f);
        }
        if (stStir.present)
        {
            using var p = BarPen(stStir.active, Color.SlateGray, 1.4f);
            for (int i = 1; i <= 11; i++)
            {
                float sx = elev.Left + i * elev.Width / 12f;
                g.DrawLine(p, sx, elev.Top + 8, sx, elev.Bottom - 8);
            }
        }

        DrawDimLine(g, elev.Left, elev.Bottom + 14, elev.Right, elev.Bottom + 14, $"L = {el.LengthM * 1000:0} mm");

        // Top Right: Live Volume & Takeoff Summary Badge
        float sumX = elev.Right + 24;
        float sumW = Math.Max(180, b.Right - sumX - 20);
        var sumRect = new RectangleF(sumX, b.Top + 16, sumW, 115);
        using (var bgBr = new SolidBrush(Color.FromArgb(240, 246, 252)))
        using (var brdrP = new Pen(Color.FromArgb(180, 205, 230), 1f))
        {
            g.FillRectangle(bgBr, sumRect);
            g.DrawRectangle(brdrP, sumRect.X, sumRect.Y, sumRect.Width, sumRect.Height);
        }

        double volBt = el.LengthM * el.WidthM * el.HeightM * (el.Qty > 0 ? el.Qty : 1);
        double areaVk = ((2 * el.HeightM) + el.WidthM) * el.LengthM * (el.Qty > 0 ? el.Qty : 1);
        double totalKg = calcs != null ? calcs.Sum(c => c.WeightKg) : 0;
        double density = volBt > 0.001 ? (totalKg / volBt) : 0;

        using (var badgeF = new Font("Segoe UI", 8.2f, FontStyle.Bold))
        using (var badgeBr = new SolidBrush(Color.FromArgb(20, 60, 110)))
        {
            g.DrawString($"CẤU KIỆN: {el.Mark} ({el.WidthM * 1000:0}x{el.HeightM * 1000:0})", badgeF, badgeBr, sumRect.Left + 8, sumRect.Top + 8);
            g.DrawString($"• Bê tông: {volBt:0.00} m³", sf, Brushes.Black, sumRect.Left + 8, sumRect.Top + 28);
            g.DrawString($"• Ván khuôn: {areaVk:0.00} m²", sf, Brushes.Black, sumRect.Left + 8, sumRect.Top + 46);
            g.DrawString($"• Cốt thép: {totalKg:0.1} kg", sf, Brushes.DarkRed, sumRect.Left + 8, sumRect.Top + 64);
            g.DrawString($"• Hàm lượng: {density:0.0} kg/m³", sf, Brushes.DarkSlateBlue, sumRect.Left + 8, sumRect.Top + 82);
        }

        // Bottom half: 3 Cross Sections (TRÁI - GIỮA - PHẢI) like Image 0
        float secY = elev.Bottom + 40;
        float availW = b.Width - 50;
        float colSpacing = availW / 3f;
        float secW = Math.Min(95, colSpacing - 30);
        float secH = Math.Min(105, b.Bottom - secY - 45);
        if (secH < 75) secH = 75;

        string[] secNames = { "TRÁI / LEFT (Gối)", "GIỮA / CEN. (Nhịp)", "PHẢI / RIGHT (Gối)" };
        string[] stirTexts = {
            stStir.present ? "a100" : "",
            stStir.present ? "a150" : "",
            stStir.present ? "a100" : ""
        };

        int topCount = 7;
        int botCount = 7;
        int sideCount = 2;
        if (stTop.present && stTop.tag.Contains("("))
        {
            var m = System.Text.RegularExpressions.Regex.Match(stTop.tag, @"(?<cnt>\d+)c");
            if (m.Success && int.TryParse(m.Groups["cnt"].Value, out int c)) topCount = Math.Clamp(c, 2, 20);
        }
        if (stBot.present && stBot.tag.Contains("("))
        {
            var m = System.Text.RegularExpressions.Regex.Match(stBot.tag, @"(?<cnt>\d+)c");
            if (m.Success && int.TryParse(m.Groups["cnt"].Value, out int c)) botCount = Math.Clamp(c, 2, 20);
        }
        if (stSide.present && stSide.tag.Contains("("))
        {
            var m = System.Text.RegularExpressions.Regex.Match(stSide.tag, @"(?<cnt>\d+)c");
            if (m.Success && int.TryParse(m.Groups["cnt"].Value, out int sc)) sideCount = Math.Clamp(sc, 2, 12);
        }

        int extraTopCount = 0;
        var rExtra = rows.FirstOrDefault(r => r.Role == "TOP-Goi");
        if (rExtra != null && int.TryParse(rExtra.OverrideQtyExpr, out int ec))
        {
            extraTopCount = Math.Clamp(ec, 1, 10);
        }

        for (int i = 0; i < 3; i++)
        {
            float cx = b.Left + 30 + (i * colSpacing) + (colSpacing - secW) / 2f;
            var sRect = new RectangleF(cx, secY, secW, secH);

            // Title above section
            g.DrawString(secNames[i], tf, tbr, sRect.Left - 4, sRect.Top - 18);

            // Concrete fill
            g.FillRectangle(new SolidBrush(ConcFill), sRect);
            g.DrawRectangle(new Pen(ConcStroke, 1.4f), sRect.X, sRect.Y, sRect.Width, sRect.Height);

            // Level indicator FL+0.000
            using (var levF = new Font("Segoe UI", 6.8f, FontStyle.Regular))
            {
                g.DrawString("FL+0.000", levF, Brushes.DarkGreen, sRect.Right - 8, sRect.Top - 12);
                g.DrawLine(Pens.DarkGreen, sRect.Right - 12, sRect.Top, sRect.Right + 8, sRect.Top);
            }

            // Outer Stirrup Hoop A
            using (var pStir = BarPen(stStir.active, Color.DarkGoldenrod, 1.5f))
            {
                g.DrawRectangle(pStir, sRect.Left + 7, sRect.Top + 7, sRect.Width - 14, sRect.Height - 14);
                // Stirrup hook A at top-left
                g.DrawLine(pStir, sRect.Left + 7, sRect.Top + 7, sRect.Left + 14, sRect.Top + 14);
            }

            // Inner Horizontal Ties B, C
            if (stTieC.present || el.HeightM >= 0.70)
            {
                using var pTieH = BarPen(stTieC.active, Color.SlateGray, 1.3f);
                int hTies = el.HeightM >= 0.90 ? 2 : 1;
                for (int ht = 1; ht <= hTies; ht++)
                {
                    float yTie = sRect.Top + ht * sRect.Height / (hTies + 1f);
                    g.DrawLine(pTieH, sRect.Left + 7, yTie, sRect.Right - 7, yTie);
                    g.DrawLine(pTieH, sRect.Left + 7, yTie, sRect.Left + 12, yTie - 4);
                    g.DrawLine(pTieH, sRect.Right - 7, yTie, sRect.Right - 12, yTie - 4);
                }
            }

            // Inner Vertical Ties D
            if (stTieC.present || el.WidthM >= 0.80)
            {
                using var pTieV = BarPen(stTieC.active, Color.DarkMagenta, 1.3f);
                int vBranches = el.WidthM >= 1.20 ? 4 : (el.WidthM >= 0.90 ? 3 : 2);
                for (int vt = 1; vt <= vBranches; vt++)
                {
                    float xTie = sRect.Left + vt * sRect.Width / (vBranches + 1f);
                    g.DrawLine(pTieV, xTie, sRect.Top + 7, xTie, sRect.Bottom - 7);
                    g.DrawLine(pTieV, xTie, sRect.Top + 7, xTie - 4, sRect.Top + 12);
                    g.DrawLine(pTieV, xTie, sRect.Bottom - 7, xTie + 4, sRect.Bottom - 12);
                }
            }

            // Top rebar dots - Layer 1
            using (var brTop = new SolidBrush(stTop.active ? BarActive : BarNormal))
            {
                for (int t = 0; t < topCount; t++)
                {
                    float dotX = sRect.Left + 10 + (t * (sRect.Width - 20) / Math.Max(1, topCount - 1));
                    g.FillEllipse(brTop, dotX - 2.5f, sRect.Top + 8, 5, 5);
                }

                // Layer 2 reinforcement at Gối Trái (i == 0) and Gối Phải (i == 2)
                if ((i == 0 || i == 2) && extraTopCount > 0)
                {
                    for (int et = 0; et < extraTopCount; et++)
                    {
                        float dotX = sRect.Left + 14 + (et * (sRect.Width - 28) / Math.Max(1, extraTopCount - 1));
                        g.FillEllipse(brTop, dotX - 2.5f, sRect.Top + 18, 5, 5);
                    }
                }
            }

            // Bottom rebar dots
            using (var brBot = new SolidBrush(stBot.active ? BarActive : BarNormal))
            {
                for (int bt = 0; bt < botCount; bt++)
                {
                    float dotX = sRect.Left + 10 + (bt * (sRect.Width - 20) / Math.Max(1, botCount - 1));
                    g.FillEllipse(brBot, dotX - 2.5f, sRect.Bottom - 13, 5, 5);
                }
            }

            // Side rebar dots
            if (stSide.present)
            {
                using var brSide = new SolidBrush(stSide.active ? BarActive : Color.Teal);
                int pairs = Math.Max(1, sideCount / 2);
                for (int p = 0; p < pairs; p++)
                {
                    float y = sRect.Top + 20 + ((p + 1) * (sRect.Height - 40) / (pairs + 1));
                    g.FillEllipse(brSide, sRect.Left + 8, y - 2.5f, 5, 5);
                    g.FillEllipse(brSide, sRect.Right - 13, y - 2.5f, 5, 5);
                }
            }

            // Dimensions b x h
            DrawDimLine(g, sRect.Left, sRect.Bottom + 12, sRect.Right, sRect.Bottom + 12, $"{el.WidthM * 1000:0}");
            if (i == 0)
            {
                DrawDimLine(g, sRect.Left - 12, sRect.Top, sRect.Left - 12, sRect.Bottom, $"{el.HeightM * 1000:0}");
            }

            // Stirrup note below
            string stirSub = $"Đai: {stirTexts[i]}";
            using var subF = new Font("Segoe UI", 7.2f, FontStyle.Bold);
            g.DrawString(stirSub, subF, Brushes.DimGray, sRect.Left - 2, sRect.Bottom + 26);
        }
    }

    private static void DrawColumnDiagram(Graphics g, Rectangle b, ElementDataDto el, IReadOnlyList<RebarConfigRowDto> rows, IReadOnlyList<RebarCalcPreview> calcs, int selIdx)
    {
        var elev = new RectangleF(b.Left + 75, b.Top + 44, 110, Math.Min(190, b.Height - 75));
        var sec = new RectangleF(elev.Right + 85, b.Top + 56, 120, 120);

        using var tf = new Font("Segoe UI", 8.5f, FontStyle.Bold);
        g.DrawString("MAT DUNG COT (NOI SOLE)", tf, new SolidBrush(ConcStroke), elev.Left - 25, b.Top + 12);
        g.DrawString("MAT CAT NGANG X x Y", tf, new SolidBrush(ConcStroke), sec.Left - 10, b.Top + 12);

        g.FillRectangle(new SolidBrush(ConcFill), elev);
        g.DrawRectangle(new Pen(ConcStroke, 1.5f), elev.X, elev.Y, elev.Width, elev.Height);
        g.FillRectangle(new SolidBrush(ConcFill), sec);
        g.DrawRectangle(new Pen(ConcStroke, 1.5f), sec.X, sec.Y, sec.Width, sec.Height);

        var stA = FindRoleState(rows, calcs, selIdx, "MAIN.A", "COL_MAIN_A", "C1A");
        var stB = FindRoleState(rows, calcs, selIdx, "MAIN.B", "COL_MAIN_B", "C1B");
        var stStir = FindRoleState(rows, calcs, selIdx, "STIR.R", "STIRRUP_RECT", "D1");
        var stIn = FindRoleState(rows, calcs, selIdx, "STIR.IN", "STIRRUP_INNER", "D2");
        var stC = FindRoleState(rows, calcs, selIdx, "STIR.C", "STIRRUP_C", "D3");

        // Stagger bars
        using (var pA = BarPen(stA.active, BarNormal, 2.6f))
        {
            g.DrawLine(pA, elev.Left + 18, elev.Bottom - 8, elev.Left + 18, elev.Top - 20);
            g.DrawLine(pA, elev.Left + 18, elev.Bottom - 8, elev.Left + 4, elev.Bottom - 8);
        }
        using (var pB = BarPen(stB.active, BarSecondary, 2.6f))
        {
            g.DrawLine(pB, elev.Right - 18, elev.Bottom - 8, elev.Right - 18, elev.Top - 4);
            g.DrawLine(pB, elev.Right - 18, elev.Bottom - 8, elev.Right - 4, elev.Bottom - 8);
        }
        if (stStir.present)
        {
            using var pS = BarPen(stStir.active, Color.SlateGray, 1.5f);
            for (int i = 1; i <= 10; i++)
            {
                float sy = elev.Top + i * elev.Height / 11f;
                g.DrawLine(pS, elev.Left + 12, sy, elev.Right - 12, sy);
            }
            g.DrawRectangle(pS, sec.Left + 10, sec.Top + 10, sec.Width - 20, sec.Height - 20);
        }
        if (stIn.present)
        {
            using var pI = BarPen(stIn.active, Color.DarkMagenta, 1.8f);
            g.DrawRectangle(pI, sec.Left + 26, sec.Top + 10, sec.Width - 52, sec.Height - 20);
        }
        if (stC.present)
        {
            using var pC = BarPen(stC.active, Color.DarkGoldenrod, 2f);
            float mx = (sec.Left + sec.Right) / 2f;
            g.DrawLine(pC, mx, sec.Top + 10, mx, sec.Bottom - 10);
        }

        DrawDimLine(g, elev.Left - 18, elev.Top, elev.Left - 18, elev.Bottom, $"H={el.HeightM * 1000:0}");
        DrawDimLine(g, sec.Left, sec.Bottom + 20, sec.Right, sec.Bottom + 20, $"X={el.LengthM * 1000:0}");
        DrawDimLine(g, sec.Right + 8, sec.Top, sec.Right + 8, sec.Bottom, $"Y={el.WidthM * 1000:0}");

        if (stA.present) DrawCallout(g, elev.Left + 18, elev.Top + 15, elev.Right + 10, b.Top + 34, stA.tag, stA.active);
        if (stB.present) DrawCallout(g, elev.Right - 18, elev.Top + 42, elev.Right + 10, b.Top + 56, stB.tag, stB.active);
        if (stStir.present) DrawCallout(g, sec.Right - 10, sec.Top + 26, sec.Right + 4, b.Bottom - 54, stStir.tag, stStir.active);
        if (stIn.present) DrawCallout(g, sec.Left + 26, sec.Top + 60, sec.Left - 10, b.Bottom - 26, stIn.tag, stIn.active, true);
    }

    private static void DrawWallDiagram(Graphics g, Rectangle b, ElementDataDto el, IReadOnlyList<RebarConfigRowDto> rows, IReadOnlyList<RebarCalcPreview> calcs, int selIdx)
    {
        var elev = new RectangleF(b.Left + 46, b.Top + 48, Math.Min(290, b.Width * 0.50f), Math.Min(175, b.Height - 80));
        var sec = new RectangleF(elev.Right + 65, elev.Top, 42, elev.Height);

        using var tf = new Font("Segoe UI", 8.5f, FontStyle.Bold);
        g.DrawString("MAT DUNG VACH L x H", tf, new SolidBrush(ConcStroke), elev.Left, b.Top + 12);
        g.DrawString("BE DAY Y", tf, new SolidBrush(ConcStroke), sec.Left - 10, b.Top + 12);

        g.FillRectangle(new SolidBrush(ConcFill), elev);
        g.DrawRectangle(new Pen(ConcStroke, 1.5f), elev.X, elev.Y, elev.Width, elev.Height);
        g.FillRectangle(new SolidBrush(ConcFill), sec);
        g.DrawRectangle(new Pen(ConcStroke, 1.5f), sec.X, sec.Y, sec.Width, sec.Height);

        var stV = FindRoleState(rows, calcs, selIdx, "MAIN.V", "WALL_VERT", "VD");
        var stH = FindRoleState(rows, calcs, selIdx, "MAIN.H", "WALL_HORZ", "HN");
        var stC = FindRoleState(rows, calcs, selIdx, "TIE.C", "WALL_C_TIE", "DC");

        using (var pV = BarPen(stV.active, BarNormal))
        {
            for (int i = 1; i <= 7; i++)
            {
                float x = elev.Left + i * elev.Width / 8f;
                g.DrawLine(pV, x, elev.Top + 6, x, elev.Bottom - 6);
            }
            g.DrawLine(pV, sec.Left + 8, sec.Top + 6, sec.Left + 8, sec.Bottom - 6);
            g.DrawLine(pV, sec.Right - 8, sec.Top + 6, sec.Right - 8, sec.Bottom - 6);
        }
        using (var pH = BarPen(stH.active, BarSecondary))
            for (int i = 1; i <= 5; i++)
            {
                float y = elev.Top + i * elev.Height / 6f;
                g.DrawLine(pH, elev.Left + 6, y, elev.Right - 6, y);
            }
        if (stC.present)
        {
            using var pC = BarPen(stC.active, Color.DarkGoldenrod, 2.2f);
            for (int i = 1; i <= 4; i++)
            {
                float y = sec.Top + i * sec.Height / 5f;
                g.DrawLine(pC, sec.Left + 8, y, sec.Right - 8, y);
            }
        }

        DrawDimLine(g, elev.Left, elev.Bottom + 20, elev.Right, elev.Bottom + 20, $"X/L = {el.LengthM * 1000:0} mm");
        DrawDimLine(g, elev.Left - 18, elev.Top, elev.Left - 18, elev.Bottom, $"H={el.HeightM * 1000:0}");
        DrawDimLine(g, sec.Left, sec.Bottom + 20, sec.Right, sec.Bottom + 20, $"Y={el.WidthM * 1000:0}");

        if (stV.present) DrawCallout(g, elev.Left + elev.Width * 0.25f, elev.Top + 18, elev.Left + 10, b.Top + 32, stV.tag, stV.active);
        if (stH.present) DrawCallout(g, elev.Left + elev.Width * 0.7f, elev.Top + elev.Height * 0.5f, elev.Right - 10, b.Bottom - 26, stH.tag, stH.active);
        if (stC.present) DrawCallout(g, sec.Right - 8, sec.Top + sec.Height * 0.4f, sec.Right + 6, b.Top + 68, stC.tag, stC.active);
    }

    private static void DrawSlabDiagram(Graphics g, Rectangle b, ElementDataDto el, IReadOnlyList<RebarConfigRowDto> rows, IReadOnlyList<RebarCalcPreview> calcs, int selIdx)
    {
        var plan = new RectangleF(b.Left + 46, b.Top + 48, Math.Min(230, b.Width * 0.44f), Math.Min(145, b.Height - 85));
        var sec = new RectangleF(plan.Right + 50, b.Top + 82, Math.Min(195, b.Width * 0.34f), 46);

        using var tf = new Font("Segoe UI", 8.5f, FontStyle.Bold);
        g.DrawString("MAT BANG THEP SAN", tf, new SolidBrush(ConcStroke), plan.Left, b.Top + 12);
        g.DrawString("MAT CAT SAN (H)", tf, new SolidBrush(ConcStroke), sec.Left, b.Top + 12);

        g.FillRectangle(new SolidBrush(ConcFill), plan);
        g.DrawRectangle(new Pen(ConcStroke, 1.5f), plan.X, plan.Y, plan.Width, plan.Height);
        g.FillRectangle(new SolidBrush(ConcFill), sec);
        g.DrawRectangle(new Pen(ConcStroke, 1.5f), sec.X, sec.Y, sec.Width, sec.Height);

        var stBX = FindRoleState(rows, calcs, selIdx, "BOT.X", "S1");
        var stBY = FindRoleState(rows, calcs, selIdx, "BOT.Y", "S2");
        var stTX = FindRoleState(rows, calcs, selIdx, "TOP.X", "HAT.X", "S3", "M1");
        var stChair = FindRoleState(rows, calcs, selIdx, "CHAIR", "CK");

        using (var pX = BarPen(stBX.active, BarNormal))
            for (int i = 1; i <= 4; i++)
            {
                float y = plan.Top + i * plan.Height / 5f;
                g.DrawLine(pX, plan.Left + 8, y, plan.Right - 8, y);
            }
        using (var pY = BarPen(stBY.active, BarSecondary))
            for (int i = 1; i <= 5; i++)
            {
                float x = plan.Left + i * plan.Width / 6f;
                g.DrawLine(pY, x, plan.Top + 8, x, plan.Bottom - 8);
            }

        using (var pSb = BarPen(stBX.active || stBY.active, BarNormal, 2.2f))
            g.DrawLine(pSb, sec.Left + 8, sec.Bottom - 8, sec.Right - 8, sec.Bottom - 8);
        if (stTX.present)
        {
            using var pSt = BarPen(stTX.active, BarSecondary, 2.2f);
            g.DrawLine(pSt, sec.Left + 8, sec.Top + 8, sec.Right - 8, sec.Top + 8);
        }
        if (stChair.present)
        {
            using var pC = BarPen(stChair.active, Color.DarkGoldenrod, 2f);
            float cx = sec.Left + sec.Width * 0.45f;
            g.DrawLines(pC, new PointF[] { new(cx - 10, sec.Bottom - 10), new(cx, sec.Bottom - 10), new(cx, sec.Top + 10), new(cx + 18, sec.Top + 10), new(cx + 18, sec.Bottom - 10), new(cx + 28, sec.Bottom - 10) });
        }

        DrawDimLine(g, plan.Left, plan.Bottom + 20, plan.Right, plan.Bottom + 20, $"X = {el.LengthM * 1000:0} mm");
        DrawDimLine(g, plan.Left - 18, plan.Top, plan.Left - 18, plan.Bottom, $"Y={el.WidthM * 1000:0}");
        DrawDimLine(g, sec.Right + 8, sec.Top, sec.Right + 8, sec.Bottom, $"t={el.HeightM * 1000:0}");

        if (stBX.present) DrawCallout(g, plan.Left + 50, plan.Top + 28, plan.Left + 10, b.Top + 32, stBX.tag, stBX.active);
        if (stTX.present) DrawCallout(g, sec.Left + 40, sec.Top + 8, sec.Left + 4, b.Top + 56, stTX.tag, stTX.active);
        if (stChair.present) DrawCallout(g, sec.Left + sec.Width * 0.52f, sec.Top + 24, sec.Left + 10, b.Bottom - 40, stChair.tag, stChair.active);
    }

    private static void DrawStairDiagram(Graphics g, Rectangle b, ElementDataDto el, IReadOnlyList<RebarConfigRowDto> rows, IReadOnlyList<RebarCalcPreview> calcs, int selIdx)
    {
        using var tf = new Font("Segoe UI", 8.5f, FontStyle.Bold);
        g.DrawString("MAT CAT BAN THANG NGHIENG", tf, new SolidBrush(ConcStroke), b.Left + 30, b.Top + 12);
        var pts = new PointF[]
        {
            new(b.Left + 60, b.Bottom - 48),
            new(b.Left + 150, b.Bottom - 48),
            new(b.Right - 140, b.Top + 55),
            new(b.Right - 55, b.Top + 55),
            new(b.Right - 55, b.Top + 85),
            new(b.Right - 130, b.Top + 85),
            new(b.Left + 165, b.Bottom - 22),
            new(b.Left + 60, b.Bottom - 22)
        };
        g.FillPolygon(new SolidBrush(ConcFill), pts);
        g.DrawPolygon(new Pen(ConcStroke, 1.5f), pts);

        var st1 = FindRoleState(rows, calcs, selIdx, "ST-01", "ST01");
        var st3 = FindRoleState(rows, calcs, selIdx, "ST-03", "ST03");
        using (var p1 = BarPen(st1.active, BarNormal, 2.5f))
            g.DrawLines(p1, new PointF[] { new(b.Left + 70, b.Bottom - 30), new(b.Left + 160, b.Bottom - 30), new(b.Right - 135, b.Top + 76), new(b.Right - 65, b.Top + 76) });
        if (st3.present)
        {
            using var p3 = BarPen(st3.active, BarSecondary, 2.2f);
            g.DrawLine(p3, b.Right - 190, b.Top + 90, b.Right - 65, b.Top + 62);
        }
        if (st1.present) DrawCallout(g, b.Left + 210, b.Bottom - 62, b.Left + 100, b.Top + 38, st1.tag, st1.active);
        if (st3.present) DrawCallout(g, b.Right - 120, b.Top + 62, b.Right - 220, b.Top + 40, st3.tag, st3.active);
    }

    public static void RenderShapePreview(Graphics g, Rectangle b, RebarConfigRowDto? row, RebarCalcPreview? calc)
    {
        g.SmoothingMode = SmoothingMode.AntiAlias;
        g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.ClearTypeGridFit;
        g.Clear(Color.White);
        if (row == null || calc == null) return;

        string sc = (row.ShapeCode ?? "D01").Trim().ToUpperInvariant();
        using var p = new Pen(BarActive, 2.8f) { LineJoin = LineJoin.Round };
        using var f = new Font("Segoe UI", 8f, FontStyle.Bold);
        using var br = new SolidBrush(Color.FromArgb(22, 54, 92));

        float padX = 28, padY = 28;
        float w = Math.Max(60, b.Width - 2 * padX);
        float h = Math.Max(40, b.Height - 2 * padY);
        float x0 = b.Left + padX, x1 = x0 + w;
        float y0 = b.Top + padY, y1 = y0 + h;
        float cy = (y0 + y1) / 2f;

        g.DrawString($"{sc}  |  Ø{(int)row.Dia}", f, br, b.Left + 8, b.Top + 6);

        if (sc == "D01")
        {
            g.DrawLine(p, x0, cy, x1, cy);
            g.DrawString($"L = {calc.MidMm:0} mm", f, br, x0 + w * 0.3f, cy - 18);
        }
        else if (sc is "D06" or "D11")
        {
            g.DrawLines(p, new PointF[] { new(x0, y0 + 10), new(x0, y1 - 10), new(x1, y1 - 10) });
            g.DrawString($"L={calc.LeftMm:0}", f, br, x0 + 4, cy - 8);
            g.DrawString($"Mid={calc.MidMm:0}", f, br, x0 + w * 0.35f, y1 - 24);
        }
        else if (sc is "D39" or "D41")
        {
            var r = new RectangleF(x0 + 20, y0 + 8, w - 40, h - 16);
            g.DrawRectangle(p, r.X, r.Y, r.Width, r.Height);
            g.DrawLine(p, r.Left, r.Top, r.Left + 14, r.Top + 14);
            g.DrawLine(p, r.Left, r.Top, r.Left + 18, r.Top + 8);
            g.DrawString($"Chu vi={calc.MidMm:0} + 2*{calc.LeftMm:0}", f, br, x0 + 12, y1 - 14);
        }
        else if (sc is "D24" or "D02" or "D04")
        {
            g.DrawLines(p, new PointF[] { new(x0 + 12, cy - 18), new(x0, cy - 18), new(x0, cy + 12), new(x1, cy + 12), new(x1, cy - 18), new(x1 - 12, cy - 18) });
            g.DrawString($"L={calc.LeftMm:0} | Mid={calc.MidMm:0} | R={calc.RightMm:0}", f, br, x0 + 6, cy + 18);
        }
        else if (sc is "D17" or "D08" or "D18" or "D15")
        {
            g.DrawLines(p, new PointF[] {
                new(x0, y0 + 8), new(x0 + w * 0.18f, y0 + 8),
                new(x0 + w * 0.32f, y1 - 10), new(x0 + w * 0.68f, y1 - 10),
                new(x0 + w * 0.82f, y0 + 8), new(x1, y0 + 8)
            });
            g.DrawString($"L={calc.LeftMm:0} | L2={calc.Le1aMm:0} | Mid={calc.MidMm:0} | R={calc.RightMm:0}", f, br, b.Left + 8, y1 - 12);
        }
        else
        {
            // Standard U-bar D07 / multi
            g.DrawLines(p, new PointF[] { new(x0, y0 + 10), new(x0, y1 - 14), new(x1, y1 - 14), new(x1, y0 + 10) });
            g.DrawString($"Left={calc.LeftMm:0}", f, br, x0 + 4, cy - 12);
            g.DrawString($"Mid={calc.MidMm:0}", f, br, x0 + w * 0.33f, y1 - 28);
            g.DrawString($"Right={calc.RightMm:0}", f, br, x1 - 65, cy - 12);
        }
    }
}
