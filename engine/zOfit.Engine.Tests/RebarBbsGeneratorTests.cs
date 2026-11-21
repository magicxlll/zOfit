using System.Drawing;
using Xunit;
using ZOfit.Engine.Generators;
using ZOfit.Engine.Models;
using ZOfit.Engine.UI;

namespace ZOfit.Engine.Tests;

public class RebarBbsGeneratorTests
{
    [Theory]
    [InlineData("FTG_2WAY", "Foundation", 2)]
    [InlineData("SLAB_1WAY", "Slab", 2)]
    [InlineData("SLAB_2WAY_1LAYER", "Slab", 2)]
    [InlineData("SLAB_2WAY_2LAYER", "SlabOnGrade", 5)]
    [InlineData("SLAB_WITH_HATS", "Slab", 5)]
    [InlineData("BEAM_FULL", "Beam", 7)]
    [InlineData("COL_STAGGER", "Column", 5)]
    [InlineData("WALL_2FACE", "Wall", 3)]
    [InlineData("STAIR_6BAR", "Stair", 3)]
    [InlineData("MANHOLE_STD", "Manhole", 2)]
    [InlineData("PILE_HEAD", "Pile", 2)]
    public void Expands_All_Smart_RebarPresets_And_Links_TKT_Formulas(string tpl, string group, int expectedBars)
    {
        var el = new ElementDataDto
        {
            DataRow = 12,
            ElementId = "EL-01",
            Building = "B49",
            Group = group,
            Mark = "M1",
            Description = "Test " + tpl,
            Qty = 4,
            LengthM = 5.0,
            WidthM = 2.0,
            HeightM = 0.8,
            RebarTemplate = tpl,
            MainXDia = 22, MainXSpec = "4",
            MainYDia = 14, MainYSpec = "@200",
            TopXDia = 20, TopXSpec = "4",
            TopYDia = 12, TopYSpec = "@200",
            StirrupDia = 10, StirrupSpec = "100/200",
            BoqBucket = "Substructure"
        };

        var gen = new RebarBbsGenerator();
        var expanded = gen.ExpandSingleElementToConfigRows(el);
        Assert.Equal(expectedBars, expanded.Count);
        Assert.All(expanded, r =>
        {
            Assert.False(string.IsNullOrWhiteSpace(r.BarDir));
            Assert.False(string.IsNullOrWhiteSpace(r.DistDir));
            Assert.StartsWith("=SPEC_Cover_", r.CoverExpr);
        });

        var tkt = gen.Generate("B49", new ProjectSpecDto(), new[] { el });
        Assert.Equal("TKT_B49", tkt.SheetName);
        Assert.Equal(expectedBars, tkt.Rows.Count);
    }

    [Fact]
    public void Explicit_RebarConfig_V3_Links_1To1_To_Transparent_Segments_And_Ignores_BlockMarkers()
    {
        var wall = new ElementDataDto { DataRow = 5, ElementId = "49-W1", Building = "B49", Group = "Wall", Mark = "W1", Qty = 2, LengthM = 6.0, WidthM = 0.3, HeightM = 4.0, RebarTemplate = "WALL_2FACE", BoqBucket = "Superstructure" };
        var beam = new ElementDataDto { DataRow = 6, ElementId = "49-B1", Building = "B49", Group = "Beam", Mark = "B1", Qty = 3, LengthM = 5.05, WidthM = 0.2, HeightM = 0.45, RebarTemplate = "BEAM_FULL", BoqBucket = "Superstructure" };

        var configs = new List<RebarConfigRowDto>
        {
            new() { ConfigRow = 5, RowKind = "BLK_HDR", ElementId = "49-W1" },
            new() { ConfigRow = 6, RowKind = "BAR", ElementId = "49-W1", BarId = "W1-VD", Role = "MAIN.V", ShapeCode = "D06", Dia = 14, LayerMult = 2, BarDir = "H", DistDir = "X", Note = "Vert" },
            new() { ConfigRow = 7, RowKind = "BAR", ElementId = "49-W1", BarId = "W1-HN", Role = "MAIN.H", ShapeCode = "D07", Dia = 12, LayerMult = 2, BarDir = "X", DistDir = "H", Note = "Horz" },
            new() { ConfigRow = 8, RowKind = "BLK_END", ElementId = "49-W1" },
            new() { ConfigRow = 9, RowKind = "BLK_HDR", ElementId = "49-B1" },
            new() { ConfigRow = 10, RowKind = "BAR", ElementId = "49-B1", BarId = "B1-ST1", Role = "STIR.R", ShapeCode = "D39", Dia = 8, LayerMult = 1, BarDir = "PERIM_YH", DistDir = "ZONE2_X", Note = "Stirrup 2-zone" },
            new() { ConfigRow = 11, RowKind = "BAR", ElementId = "49-B1", BarId = "B1-SP", Role = "CUSTOM", ShapeCode = "D17", Dia = 16, LayerMult = 1, BarDir = "CUSTOM", DistDir = "FIXED", Note = "Custom Crank" },
            new() { ConfigRow = 12, RowKind = "BLK_END", ElementId = "49-B1" }
        };

        var tkt = new RebarBbsGenerator().Generate("B49", new ProjectSpecDto(), new[] { wall, beam }, null, configs);
        Assert.Equal(4, tkt.Rows.Count); // BLK_HDR and BLK_END ignored!

        // W1-VD (D06, row 6): Dia=L6, QtyPerEl=M6*AG6, Segs reference Z6/AC6/AF6
        Assert.Equal("=QS_RebarConfig!L6", tkt.Rows[0].Values[14]?.ToString());
        Assert.Equal("=QS_RebarConfig!M6*QS_RebarConfig!AG6", tkt.Rows[0].Values[15]?.ToString());
        Assert.Contains("QS_RebarConfig!Z6", tkt.Rows[0].Values[5]?.ToString());
        Assert.Contains("QS_RebarConfig!AC6", tkt.Rows[0].Values[6]?.ToString());

        // W1-HN (D07, row 7): F=Z7, G=AC7, H=AF7
        Assert.Equal("=QS_RebarConfig!Z7", tkt.Rows[1].Values[5]?.ToString());
        Assert.Equal("=QS_RebarConfig!AC7", tkt.Rows[1].Values[6]?.ToString());
        Assert.Equal("=QS_RebarConfig!AF7", tkt.Rows[1].Values[7]?.ToString());

        // B1-ST1 (D39, row 10): links G10/H10/Q10/V10 and hook M=Z10
        Assert.Equal("D39", tkt.Rows[2].Values[3]);
        Assert.Contains("QS_RebarConfig!V10", tkt.Rows[2].Values[5]?.ToString());
        Assert.Equal("=QS_RebarConfig!Z10", tkt.Rows[2].Values[12]?.ToString());

        // Custom D17 (row 11): links Z11, AA11, AC11, AF11
        Assert.Equal("D17", tkt.Rows[3].Values[3]);
        Assert.Equal("=QS_RebarConfig!Z11", tkt.Rows[3].Values[5]?.ToString());
        Assert.Contains("QS_RebarConfig!AA11", tkt.Rows[3].Values[6]?.ToString());
        Assert.Equal("=QS_RebarConfig!AF11", tkt.Rows[3].Values[8]?.ToString());
    }

    [Fact]
    public void VisualRebarEditor_And_ParamCalculator_Support_DiagramRender_QuickRoles_And_LiveMath()
    {
        var spec = new ProjectSpecDto();
        var beam = new ElementDataDto
        {
            DataRow = 6, ElementId = "49-GB-GBX1", Building = "B49", Group = "GroundBeam",
            Mark = "GBX1", Description = "Ground Beam GBX1", Qty = 2,
            LengthM = 5.05, WidthM = 0.2, HeightM = 0.45, RebarTemplate = "BEAM_FULL", BoqBucket = "Substructure"
        };

        var req = new EditRebarUiRequest { Spec = spec, Element = beam };
        using var form = new RebarEditorForm(req);
        Assert.Equal(7, form.CurrentBars.Count); // BEAM_FULL preloaded

        // Render element diagram & shape diagram off-screen to verify zero GDI+ exceptions across all groups
        using (var bmp = new Bitmap(640, 320))
        using (var g = Graphics.FromImage(bmp))
        {
            var calcs = form.CurrentBars.Select(b => RebarParamCalculator.ComputeRow(spec, beam, b)).ToList();
            StructuralDiagramRenderer.RenderElementDiagram(g, new Rectangle(0, 0, 640, 320), beam, form.CurrentBars, calcs, 0);
            StructuralDiagramRenderer.RenderShapePreview(g, new Rectangle(0, 0, 220, 140), form.CurrentBars[0], calcs[0]);
        }

        // Verify stirrup 2-zone math matches AAA_BOQ_CAL 1G1 exact example (L=5050, @1=150, @2=200 => Qty=32, Bv=40 => Mid=980 if Beam)
        var stir = RebarParamCalculator.CreateQuickRoleRow(beam, spec, "STIR.BEAM", 9);
        stir.CoverExpr = "40"; stir.Dia = 8;
        var stirCalc = RebarParamCalculator.ComputeRow(spec, beam, stir);
        Assert.Equal(980, stirCalc.MidMm);
        Assert.Equal(32, stirCalc.QtyBar);
        Assert.Equal(1140, stirCalc.LbarMm);

        // Unlimited bar adding + Accept
        form.AddQuickRole("CUSTOM");
        Assert.Equal(8, form.CurrentBars.Count);
        form.AcceptResult();
        Assert.True(form.Result.Accepted);
        Assert.Equal(8, form.Result.Rows.Count);
    }
}
