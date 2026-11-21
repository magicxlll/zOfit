using Xunit;
using ZOfit.Engine.Models;
using ZOfit.Engine.Parsers;

namespace ZOfit.Engine.Tests;

public class CadBeamSectionParserTests
{
    [Fact]
    public void Tcvn3Converter_ConvertsAccurately()
    {
        string raw = "CHI TIÕT DÇM TÇNG 1 - THÐP TR£N - THÐP D¦íI - THÐP §AI";
        string converted = Tcvn3Converter.ToUnicode(raw);

        Assert.Contains("THÉP TRÊN", converted);
        Assert.Contains("THÉP DƯỚI", converted);
        Assert.Contains("THÉP ĐAI", converted);
    }

    [Fact]
    public void CadBeamSectionParser_ParsesImage0BeamTableData()
    {
        // Replicate the real CAD entities extracted from Image 0 (Drawing: S05-WATER TREATMENT.dwg)
        var entities = new List<CadEntityItemDto>
        {
            // Title
            new() { Handle = "46240D", Type = "TEXT", Text = "CHI TIẾT DẦM TẦNG 1 / 1ST FLOOR BEAM DETAIL", Position = new[] { 3147116.1, -780963.2 } },
            // Row 1: Beam Name Blocks (HA-ElementName)
            new() { Handle = "464D42", Type = "TEXT", Text = "T£N/ NAME", Position = new[] { 3147283.9, -781658.7 } },
            new()
            {
                Handle = "464D78", Type = "INSERT", BlockName = "HA-ElementName", Position = new[] { 3152053.6, -781644.4 },
                Attributes = new Dictionary<string, string> { { "TENCK", "1MBX1" }, { "KT", "(800x800)" }, { "LEN", "" } }
            },
            new()
            {
                Handle = "46534B", Type = "INSERT", BlockName = "HA-ElementName", Position = new[] { 3153110.1, -781644.4 },
                Attributes = new Dictionary<string, string> { { "TENCK", "1MBX9" }, { "KT", "(800x800)" }, { "LEN", "" } }
            },
            // Row 2: Location Columns
            new() { Handle = "464D43", Type = "TEXT", Text = "VÞ TRÝ/ LOCATION", Position = new[] { 3147133.9, -781958.7 } },
            new() { Handle = "464D61", Type = "TEXT", Text = "TR¸I/ LEFT", Position = new[] { 3149298.0, -781958.7 } },
            new() { Handle = "464D5A", Type = "TEXT", Text = "GI÷A/ CEN.", Position = new[] { 3152983.3, -781958.7 } },
            new() { Handle = "464D5B", Type = "TEXT", Text = "PH¶I/ RIGHT", Position = new[] { 3156161.0, -781958.7 } },
            // Row 3: Top Bar
            new() { Handle = "464D45", Type = "TEXT", Text = "THÐP TR£N/ TOP BAR", Position = new[] { 3147010.5, -784258.7 } },
            new() { Handle = "464D57", Type = "TEXT", Text = "7 D22", Position = new[] { 3149421.3, -784258.7 } },
            new() { Handle = "465360", Type = "TEXT", Text = "7 D22", Position = new[] { 3153070.3, -784258.7 } },
            new() { Handle = "4653AA", Type = "TEXT", Text = "7 D22", Position = new[] { 3156242.8, -784258.7 } },
            // Row 4: Bottom Bar
            new() { Handle = "464D46", Type = "TEXT", Text = "THÐP D¦íI/ BOTTOM BAR", Position = new[] { 3146898.5, -784558.7 } },
            new() { Handle = "464D58", Type = "TEXT", Text = "7 D20", Position = new[] { 3149421.3, -784558.7 } },
            new() { Handle = "465361", Type = "TEXT", Text = "7 D20", Position = new[] { 3153070.3, -784558.7 } },
            new() { Handle = "4653AB", Type = "TEXT", Text = "7 D20", Position = new[] { 3156242.8, -784558.7 } },
            // Row 5: Stirrup
            new() { Handle = "464D49", Type = "TEXT", Text = "THÐP §AI/ STIRRUP", Position = new[] { 3147070.5, -784858.7 } },
            new() { Handle = "464D64", Type = "TEXT", Text = "A: D12a100", Position = new[] { 3148349.0, -784858.7 } },
            new() { Handle = "464D65", Type = "TEXT", Text = "B,C: D12a100", Position = new[] { 3150762.4, -784858.7 } },
            new() { Handle = "464D6C", Type = "TEXT", Text = "A: D12a150", Position = new[] { 3152008.0, -784858.7 } },
            new() { Handle = "464D6D", Type = "TEXT", Text = "B,C: D12a150", Position = new[] { 3153541.9, -784858.7 } },
            new() { Handle = "464D6E", Type = "TEXT", Text = "A: D12a100", Position = new[] { 3155586.0, -784858.7 } },
            new() { Handle = "464D6F", Type = "TEXT", Text = "B,C: D12a100", Position = new[] { 3156752.9, -784858.7 } },
            // Row 6: Side Bar
            new() { Handle = "464DE5", Type = "TEXT", Text = "THÐP B£N/ SIDE BAR", Position = new[] { 3147022.2, -785150.2 } },
            new() { Handle = "464D85", Type = "TEXT", Text = "2 D18", Position = new[] { 3149431.0, -785158.7 } },
            new() { Handle = "466BFD", Type = "TEXT", Text = "2 D18", Position = new[] { 3152973.0, -785158.7 } },
            new() { Handle = "466BFE", Type = "TEXT", Text = "2 D18", Position = new[] { 3156228.5, -785158.7 } }
        };

        var parser = new CadBeamSectionParser();
        var resp = parser.Parse(new CadBeamSectionParseRequest
        {
            Building = "BLD",
            Floor = "TANG1",
            DefaultLengthM = 6.0,
            Entities = entities
        });

        Assert.Equal(2, resp.BeamCount);
        Assert.Equal(2, resp.Beams.Count);

        // Check Beam 1: 1MBX1
        var b1 = resp.Beams[0];
        Assert.Equal("1MBX1", b1.Element.Mark);
        Assert.Equal("BEAM", b1.Element.Group);
        Assert.Equal(0.80, b1.Element.WidthM);
        Assert.Equal(0.80, b1.Element.HeightM);
        Assert.Equal(6.0, b1.Element.LengthM);

        // Rebar rows for Beam 1
        Assert.True(b1.RebarRows.Count >= 4);

        var topRow = b1.RebarRows.FirstOrDefault(r => r.Role == "TOP");
        Assert.NotNull(topRow);
        Assert.Equal("D07", topRow.ShapeCode);
        Assert.Equal(22, topRow.Dia);
        Assert.Equal("7", topRow.OverrideQtyExpr);

        var botRow = b1.RebarRows.FirstOrDefault(r => r.Role == "BOT");
        Assert.NotNull(botRow);
        Assert.Equal("D07", botRow.ShapeCode);
        Assert.Equal(20, botRow.Dia);
        Assert.Equal("7", botRow.OverrideQtyExpr);

        var sideRow = b1.RebarRows.FirstOrDefault(r => r.Role == "SIDE");
        Assert.NotNull(sideRow);
        Assert.Equal("D01", sideRow.ShapeCode);
        Assert.Equal(18, sideRow.Dia);
        Assert.Equal("2", sideRow.OverrideQtyExpr);

        var stirRow = b1.RebarRows.FirstOrDefault(r => r.Role == "STIR.R");
        Assert.NotNull(stirRow);
        Assert.Equal("D39", stirRow.ShapeCode);
        Assert.Equal(12, stirRow.Dia);
        Assert.Equal(100, stirRow.Space1Mm);
        Assert.Equal(150, stirRow.Space2Mm);

        // Check Beam 2: 1MBX9
        var b2 = resp.Beams[1];
        Assert.Equal("1MBX9", b2.Element.Mark);
        Assert.Equal(0.80, b2.Element.WidthM);
        Assert.Equal(0.80, b2.Element.HeightM);
    }

    [Fact]
    public void CadBeamSectionParser_ParsesWideBeamWithMultiLayerAndTies()
    {
        // Test parsing the 1300x1000 beam from user's uploaded Image 0
        // (1MBX3, 1MBX5, 1MBX7) with 13 D25 + 5 D25, B,C,D: D12a100, D25a1000, 4 D18
        var entities = new List<CadEntityItemDto>
        {
            // Row 1: Beam Name Blocks (HA-ElementName)
            new() { Handle = "101", Type = "TEXT", Text = "TÊN/ NAME", Position = new[] { 100.0, 1000.0 } },
            new() { Handle = "102", Type = "TEXT", Text = "1MBX3 (1300X1000) 1MBX5 (1300X1000) 1MBX7 (1300X1000)", Position = new[] { 300.0, 1000.0 } },
            // Row 2: Location Columns
            new() { Handle = "103", Type = "TEXT", Text = "VỊ TRÍ/ LOCATION", Position = new[] { 100.0, 900.0 } },
            new() { Handle = "104", Type = "TEXT", Text = "TRÁI/ LEFT", Position = new[] { 250.0, 900.0 } },
            new() { Handle = "105", Type = "TEXT", Text = "GIỮA/ CEN.", Position = new[] { 500.0, 900.0 } },
            new() { Handle = "106", Type = "TEXT", Text = "PHẢI/ RIGHT", Position = new[] { 750.0, 900.0 } },
            // Row 3: Top Bar
            new() { Handle = "107", Type = "TEXT", Text = "THÉP TRÊN/ TOP BAR", Position = new[] { 100.0, 700.0 } },
            new() { Handle = "108", Type = "TEXT", Text = "13 D25 + 5 D25", Position = new[] { 250.0, 700.0 } },
            new() { Handle = "109", Type = "TEXT", Text = "13 D25", Position = new[] { 500.0, 700.0 } },
            new() { Handle = "110", Type = "TEXT", Text = "13 D25 + 5 D25", Position = new[] { 750.0, 700.0 } },
            // Row 4: Bottom Bar
            new() { Handle = "111", Type = "TEXT", Text = "THÉP DƯỚI/ BOTTOM BAR", Position = new[] { 100.0, 600.0 } },
            new() { Handle = "112", Type = "TEXT", Text = "13 D25", Position = new[] { 250.0, 600.0 } },
            new() { Handle = "113", Type = "TEXT", Text = "13 D25", Position = new[] { 500.0, 600.0 } },
            new() { Handle = "114", Type = "TEXT", Text = "13 D25", Position = new[] { 750.0, 600.0 } },
            // Row 5: Stirrup
            new() { Handle = "115", Type = "TEXT", Text = "THÉP ĐAI/ STIRRUP", Position = new[] { 100.0, 500.0 } },
            new() { Handle = "116", Type = "TEXT", Text = "A: D12a100", Position = new[] { 200.0, 500.0 } },
            new() { Handle = "117", Type = "TEXT", Text = "B,C,D: D12a100", Position = new[] { 350.0, 500.0 } },
            new() { Handle = "118", Type = "TEXT", Text = "A: D12a100", Position = new[] { 480.0, 500.0 } },
            new() { Handle = "119", Type = "TEXT", Text = "B,C,D: D12a100", Position = new[] { 600.0, 500.0 } },
            new() { Handle = "120", Type = "TEXT", Text = "A: D12a100", Position = new[] { 720.0, 500.0 } },
            new() { Handle = "121", Type = "TEXT", Text = "B,C,D: D12a100", Position = new[] { 850.0, 500.0 } },
            // Row 6: Support Bar
            new() { Handle = "122", Type = "TEXT", Text = "THÉP GIÁ/ SUPPORT BAR", Position = new[] { 100.0, 400.0 } },
            new() { Handle = "123", Type = "TEXT", Text = "D25a1000", Position = new[] { 250.0, 400.0 } },
            new() { Handle = "124", Type = "TEXT", Text = "D25a1000", Position = new[] { 750.0, 400.0 } },
            // Row 7: Side Bar
            new() { Handle = "125", Type = "TEXT", Text = "THÉP BÊN/ SIDE BAR", Position = new[] { 100.0, 300.0 } },
            new() { Handle = "126", Type = "TEXT", Text = "4 D18", Position = new[] { 500.0, 300.0 } }
        };

        var parser = new CadBeamSectionParser();
        var resp = parser.Parse(new CadBeamSectionParseRequest
        {
            Building = "BLD",
            Floor = "TANG1",
            DefaultLengthM = 6.0,
            Entities = entities
        });

        Assert.Equal(3, resp.BeamCount);
        var b = resp.Beams[0];
        Assert.Equal("1MBX3", b.Element.Mark);
        Assert.Equal(1.30, b.Element.WidthM);
        Assert.Equal(1.00, b.Element.HeightM);

        // Check Top bars: T1 (suốt, D07, 13D25) and T_GT (gia cường, D06, 5D25)
        var t1 = b.RebarRows.FirstOrDefault(r => r.BarId == "T1");
        Assert.NotNull(t1);
        Assert.Equal("D07", t1.ShapeCode);
        Assert.Equal(25, t1.Dia);
        Assert.Equal("13", t1.OverrideQtyExpr);

        var tgt = b.RebarRows.FirstOrDefault(r => r.BarId == "T_GT");
        Assert.NotNull(tgt);
        Assert.Equal("D06", tgt.ShapeCode);
        Assert.Equal(25, tgt.Dia);
        Assert.Equal("5", tgt.OverrideQtyExpr);

        // Check Bot bars: B1 (suốt, D07, 13D25)
        var b1 = b.RebarRows.FirstOrDefault(r => r.BarId == "B1");
        Assert.NotNull(b1);
        Assert.Equal("D07", b1.ShapeCode);
        Assert.Equal(25, b1.Dia);
        Assert.Equal("13", b1.OverrideQtyExpr);

        // Check Side bar: S1 (4 D18)
        var s1 = b.RebarRows.FirstOrDefault(r => r.BarId == "S1");
        Assert.NotNull(s1);
        Assert.Equal("D01", s1.ShapeCode);
        Assert.Equal(18, s1.Dia);
        Assert.Equal("4", s1.OverrideQtyExpr);

        // Check Stirrup A: DA (D39, D12a100)
        var da = b.RebarRows.FirstOrDefault(r => r.BarId == "DA");
        Assert.NotNull(da);
        Assert.Equal("D39", da.ShapeCode);
        Assert.Equal(12, da.Dia);
        Assert.Equal(100, da.Space1Mm);

        // Check Horizontal ties B, C: DBC (D24, layerMult=2)
        var dbc = b.RebarRows.FirstOrDefault(r => r.BarId == "DBC");
        Assert.NotNull(dbc);
        Assert.Equal("D24", dbc.ShapeCode);
        Assert.Equal(2, dbc.LayerMult);

        // Check Vertical ties D: DD (D24, layerMult=4)
        var dd = b.RebarRows.FirstOrDefault(r => r.BarId == "DD");
        Assert.NotNull(dd);
        Assert.Equal("D24", dd.ShapeCode);
        Assert.Equal(4, dd.LayerMult);

        // Check Support bar: G1
        var g1 = b.RebarRows.FirstOrDefault(r => r.BarId == "G1");
        Assert.NotNull(g1);
        Assert.Equal(25, g1.Dia);
        Assert.Equal(1000, g1.Space1Mm);
    }
}
