using Xunit;
using ZOfit.Engine.Generators;
using ZOfit.Engine.Models;

namespace ZOfit.Engine.Tests;

public class RcDetailGeneratorTests
{
    private static List<ElementDataDto> SampleElements() => new()
    {
        new() { DataRow = 5, ElementId = "49-FTG-PC1", Building = "49.ELECTRIC SUPPLY", Group = "Foundation", Mark = "PC1", Description = "Footing PC1", Qty = 10, LengthM = 1.0, WidthM = 1.0, HeightM = 1.0, TopElevM = 19.6, BotElevM = 18.6, IsBuried = "Y", RebarTemplate = "FTG_2WAY", MainXDia = 14, MainXSpec = "@150", MainYDia = 14, MainYSpec = "@150", BoqBucket = "Substructure" },
        new() { DataRow = 6, ElementId = "49-GB-GBX1", Building = "49.ELECTRIC SUPPLY", Group = "GroundBeam", Mark = "GBX1", Description = "Ground Beam GBX1", Qty = 2, LengthM = 48.6, WidthM = 0.6, HeightM = 1.0, TopElevM = 19.6, BotElevM = 18.6, IsBuried = "Y", RebarTemplate = "BEAM_FULL", MainXDia = 22, MainXSpec = "4", TopXDia = 22, TopXSpec = "4", StirrupDia = 10, StirrupSpec = "100/200", BoqBucket = "Substructure" },
        new() { DataRow = 7, ElementId = "49-SLB-GS1", Building = "49.ELECTRIC SUPPLY", Group = "SlabOnGrade", Mark = "GS1", Description = "Ground Slab GS1", Qty = 1, LengthM = 40.0, WidthM = 21.1, HeightM = 0.25, TopElevM = 19.7, BotElevM = 19.45, IsBuried = "N", RebarTemplate = "SLAB_2WAY_2LAYER", MainXDia = 12, MainXSpec = "@200", MainYDia = 12, MainYSpec = "@200", TopXDia = 12, TopXSpec = "@200", TopYDia = 12, TopYSpec = "@200", BoqBucket = "Substructure" },
        new() { DataRow = 8, ElementId = "49-COL-C1", Building = "49.ELECTRIC SUPPLY", Group = "Column", Mark = "C1", Description = "Column C1", Qty = 12, LengthM = 0.6, WidthM = 0.6, HeightM = 4.5, TopElevM = 24.5, BotElevM = 20.0, IsBuried = "N", RebarTemplate = "COL_STAGGER", MainXDia = 20, MainXSpec = "8", StirrupDia = 10, StirrupSpec = "100/200", BoqBucket = "Superstructure" },
        new() { DataRow = 9, ElementId = "49-PILE-P1", Building = "49.ELECTRIC SUPPLY", Group = "Pile", Mark = "PHC-D400", Description = "PHC Pile D400", Qty = 50, LengthM = 24.0, WidthM = 0.4, HeightM = 0.4, TopElevM = 19.7, BotElevM = 18.6, IsBuried = "N", RebarTemplate = "PILE_HEAD", MainXDia = 16, MainXSpec = "6", StirrupDia = 8, StirrupSpec = "@150", BoqBucket = "Substructure" }
    };

    [Fact]
    public void Generates_PerType_Lean_Concrete_With_PerElement_Override_Formulas()
    {
        var gen = new RcDetailGenerator();
        var sheet = gen.Generate("49.ELECTRIC SUPPLY", new ProjectSpecDto(), SampleElements());

        Assert.Equal("49.ELECTRIC SUPPLY_Detail", sheet.SheetName);

        var leanFtg = sheet.Rows.Single(r => r.TagId == "49-FTG-PC1::CONC_LEAN");
        Assert.Contains("SPEC_Lean_Ftg_Ext", leanFtg.Values[5]?.ToString());
        Assert.Contains("SPEC_Lean_Ftg_Thk", leanFtg.Values[7]?.ToString());
        Assert.Contains("QS_Data!P5", leanFtg.Values[7]?.ToString());

        var leanGb = sheet.Rows.Single(r => r.TagId == "49-GB-GBX1::CONC_LEAN");
        Assert.Contains("SPEC_Lean_Beam_Ext", leanGb.Values[5]?.ToString());
        Assert.Contains("SPEC_Lean_Beam_Thk", leanGb.Values[7]?.ToString());
        Assert.Contains("QS_Data!Q6", leanGb.Values[5]?.ToString());

        var leanSlab = sheet.Rows.Single(r => r.TagId == "49-SLB-GS1::CONC_LEAN");
        Assert.Contains("SPEC_Lean_Slab_Thk", leanSlab.Values[7]?.ToString());
        Assert.Contains("SPEC_Lean_Slab_Ext", leanSlab.Values[5]?.ToString());

        foreach (var row in sheet.Rows)
        {
            foreach (var cell in row.Values)
            {
                if (cell is string s && s.StartsWith('='))
                {
                    Assert.DoesNotContain("_xlfn.", s, StringComparison.OrdinalIgnoreCase);
                    Assert.DoesNotContain("#REF!", s, StringComparison.OrdinalIgnoreCase);
                }
            }
        }
    }
}
