using Xunit;
using ZOfit.Engine.Generators;
using ZOfit.Engine.Models;

namespace ZOfit.Engine.Tests;

public class FinishingAndInfraTests
{
    [Fact]
    public void FinishingEngine_Links_Doors_Via_IndexMatch_And_Chains_Roof()
    {
        var req = new GenerateFinishingRequest
        {
            Spec = new ProjectSpecDto(),
            Doors = new List<DoorWinDto>
            {
                new() { DataRow = 5, Code = "1SD1", DoorType = "DOOR", WidthM = 1.2, HeightM = 2.2 },
                new() { DataRow = 6, Code = "WD1", DoorType = "WINDOW", WidthM = 1.8, HeightM = 1.5 }
            },
            Rooms = new List<RoomFinishDto>
            {
                new() { DataRow = 5, RoomId = "RM-ROOF", Building = "Factory", Floor = "Roof", RoomName = "Main Roof R1", AreaM2 = 0, PerimM = 0, WallType = "NONE", RoofAreaM2 = 864, ParapetLenM = 124.7, ParapetHeightM = 1.0 },
                new() { DataRow = 6, RoomId = "RM-WC1", Building = "Office", Floor = "1F", RoomName = "Toilet Men", AreaM2 = 25, PerimM = 20, HeightM = 3.2, WallType = "B200", WallLenM = 20, DoorCodes = "1SD1*1,WD1*2", FloorFinish = "TILE", WallTileHeightM = 2.4, CeilingType = "GYPSUM", IsWetArea = "Y" }
            }
        };

        var sheet = new FinishingEngine().Generate(req);
        Assert.Equal("Finishing_Detail", sheet.SheetName);

        var masonry = sheet.Rows.Single(r => r.TagId == "RM-WC1::MASONRY");
        Assert.Contains("INDEX(QS_DoorWin!$E$5:$E$500,MATCH(\"1SD1\",QS_DoorWin!$A$5:$A$500,0))", masonry.Values[11]?.ToString());

        var wetWp = sheet.Rows.Single(r => r.TagId == "RM-WC1::WET_WP");
        Assert.Contains("SPEC_WP_CurbH", wetWp.Values[7]?.ToString());

        var slope = sheet.Rows.Single(r => r.TagId == "FIN_ROOF_SLOPE");
        Assert.StartsWith("=L", slope.Values[5]?.ToString()); // chained to base Roof WP2 row
    }

    [Fact]
    public void InfraEngine_Generates_Pipeline_Road_Fence_With_Displacement()
    {
        var req = new GenerateInfraRequest
        {
            Spec = new ProjectSpecDto(),
            Items = new List<InfraItemDto>
            {
                new() { DataRow = 5, InfraId = "INF-P1", Category = "PIPE", Code = "D400", Description = "Storm Drain D400", Qty = 1, LengthM = 120, DepthOrThkM = 1.8, DiaMm = 400 },
                new() { DataRow = 6, InfraId = "INF-R1", Category = "ROAD_BTXM", Code = "RD-A", Description = "Internal Road 7m", Qty = 1, LengthM = 100, WidthM = 7.0, DepthOrThkM = 0.22, SubBaseT1M = 0.15, ExtraSidewalkWM = 1.5, CurbLenM = 200, ParamA = 5.0 },
                new() { DataRow = 7, InfraId = "INF-F1", Category = "FENCE", Code = "F-1", Description = "Perimeter Fence", Qty = 1, LengthM = 300, DepthOrThkM = 2.4, ParamA = 0.8, ParamB = 3.0 }
            }
        };

        var sheet = new InfraEngine().Generate(req);
        Assert.Equal("Civil_Detail", sheet.SheetName);

        var bk = sheet.Rows.Single(r => r.TagId == "INF-P1::BK");
        Assert.Contains("PI()", bk.Values[11]?.ToString()); // Pipe displacement deduction

        var dowel = sheet.Rows.Single(r => r.TagId == "INF-R1::DOWEL");
        Assert.Contains("SPEC_Road_DowelSpace", dowel.Values[4]?.ToString());

        var fenceBrick = sheet.Rows.Single(r => r.TagId == "INF-F1::BRICK");
        Assert.NotNull(fenceBrick);
    }
}
