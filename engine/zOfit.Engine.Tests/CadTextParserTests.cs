using Xunit;
using ZOfit.Engine.Models;
using ZOfit.Engine.Parsers;

namespace ZOfit.Engine.Tests;

public class CadTextParserTests
{
    [Fact]
    public void Parses_All_BHFLEX_Data_Sheet_Formats_Correctly()
    {
        var req = new CadParseRequest
        {
            RawRows = new List<CadRawRowDto>
            {
                new() { SourceRow = 14, ColA = "GIA TRI" }, // header -> skipped
                new() { SourceRow = 15, ColA = "A-W300- i0.2%-L=13m" },
                new() { SourceRow = 36, ColA = "B-W300- L=6m" },
                new() { SourceRow = 40, ColA = "Box:600x600-L11m-i0.2%" },
                new() { SourceRow = 43, ColA = "Box:800x800-L10m-i0.2%" },
                new() { SourceRow = 45, ColA = "D200-L44m- i0.5%" },
                new() { SourceRow = 46, ColA = "D200-L4m" },
                new() { SourceRow = 90, ColA = "TN-TENGA2", ColB = "19.50", ColC = "18.10", ColD = "1" },
                new() { SourceRow = 99, ColA = "UNPARSEABLE CAD NOTE" }
            }
        };

        var res = new CadTextParser().Parse(req);
        Assert.Equal(8, res.TotalInput);
        Assert.Equal(7, res.ParsedCount);
        Assert.Equal(1, res.UnknownCount);

        var a300 = res.Items.Single(i => i.SourceRow == 15);
        Assert.Equal("TRENCH", a300.ItemType);
        Assert.Equal("A-W300", a300.CategoryKey);
        Assert.Equal(300, a300.WidthMm);
        Assert.Equal(0.2, a300.SlopePct);
        Assert.Equal(13, a300.LengthM);

        var b300 = res.Items.Single(i => i.SourceRow == 36);
        Assert.Equal("B-W300", b300.CategoryKey);
        Assert.Equal(6, b300.LengthM);

        var box = res.Items.Single(i => i.SourceRow == 40);
        Assert.Equal("BOX_CULVERT", box.ItemType);
        Assert.Equal("Box:600x600", box.CategoryKey);
        Assert.Equal(600, box.WidthMm);
        Assert.Equal(600, box.HeightMm);
        Assert.Equal(11, box.LengthM);

        var pipe = res.Items.Single(i => i.SourceRow == 45);
        Assert.Equal("PIPE", pipe.ItemType);
        Assert.Equal("D200", pipe.CategoryKey);
        Assert.Equal(200, pipe.DiameterMm);
        Assert.Equal(44, pipe.LengthM);

        var mh = res.Items.Single(i => i.SourceRow == 90);
        Assert.Equal("MANHOLE", mh.ItemType);
        Assert.Equal(1.4, mh.DepthM);
    }
}
