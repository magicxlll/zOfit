using Xunit;
using ZOfit.Engine.Models;
using ZOfit.Engine.Sync;

namespace ZOfit.Engine.Tests;

public class SyncEngineTests
{
    [Fact]
    public void Preserves_Manual_Rows_Updates_Existing_Adds_New_And_Flags_Removed()
    {
        var req = new GenerateStructureRequest
        {
            TargetBuilding = "B49",
            Spec = new ProjectSpecDto(),
            Elements = new List<ElementDataDto>
            {
                new() { DataRow = 5, ElementId = "B49-FTG-PC1", Building = "B49", Group = "Foundation", Mark = "PC1", Description = "Footing PC1", Qty = 10, LengthM = 1.2, WidthM = 1.2, HeightM = 1.0, IsBuried = "Y", RebarTemplate = "FTG_2WAY", MainXDia = 14, MainXSpec = "@150", BoqBucket = "Substructure" },
                new() { DataRow = 6, ElementId = "B49-FTG-PC7", Building = "B49", Group = "Foundation", Mark = "PC7", Description = "New Footing PC7", Qty = 3, LengthM = 2.0, WidthM = 2.0, HeightM = 1.0, IsBuried = "Y", RebarTemplate = "FTG_2WAY", MainXDia = 16, MainXSpec = "@150", BoqBucket = "Substructure" }
            },
            ExistingDetailRows = new List<ExistingDetailRowDto>
            {
                new() { RowIndex = 20, Code = "PC1", Description = "Footing PC1", Unit = "m3", TagId = "B49-FTG-PC1::CONC_SUB", SectionKey = "CONC_SUB" },
                new() { RowIndex = 21, Code = "CUSTOM-1", Description = "Special Machine Pit (DxRxC Manual)", SpecText = "Manual", Unit = "m3", NoExpr = "1", LExpr = "2.5", WExpr = "1.8", HExpr = "0.5", QtyExpr = "=PRODUCT(E21:K21)", TagId = "MANUAL", SectionKey = "CONC_SUB" },
                new() { RowIndex = 22, Code = "PC2", Description = "Old Footing PC2", Unit = "m3", NoExpr = "2", LExpr = "2", WExpr = "1", HExpr = "1", TagId = "B49-FTG-PC2::CONC_SUB", SectionKey = "CONC_SUB" }
            },
            ExistingTktRows = new List<ExistingTktRowDto>
            {
                new() { RowIndex = 11, CompName = "PC1", BarMark = "1", ShapeCode = "D07", Dia = 14, TagId = "B49-FTG-PC1::BOT_X", BoqBucket = "Substructure" },
                new() { RowIndex = 15, CompName = "Manual Anchor Plate Bar", BarMark = "99", ShapeCode = "D01", L1Expr = "1200", Dia = 12, QtyPerElExpr = "4", ElQtyExpr = "2", TagId = "MANUAL", BoqBucket = "Substructure" }
            }
        };

        var res = new SyncEngine().GenerateAndSync(req);
        var det = Assert.Single(res.DetailSheets);
        var tkt = Assert.Single(res.TktSheets);

        Assert.Equal(1, det.AddedCount);           // PC7 added
        Assert.Equal(1, det.UpdatedCount);         // PC1 updated
        Assert.Equal(1, det.PreservedManualCount); // CUSTOM-1 preserved
        Assert.Equal(1, det.RemovedFlaggedCount);  // PC2 flagged [REMOVED]

        Assert.Contains(det.Rows, r => r.RowKind == "MANUAL" && r.Values[1]?.ToString() == "Special Machine Pit (DxRxC Manual)");
        Assert.Contains(det.Rows, r => r.TagId == "B49-FTG-PC7::CONC_SUB");
        Assert.Contains(det.Rows, r => r.RowKind == "REMOVED" && r.Values[1]?.ToString()!.Contains("[REMOVED]") == true);

        Assert.Equal(1, tkt.PreservedManualCount);
        Assert.Contains(tkt.Rows, r => r.RowKind == "MANUAL" && r.Values[1]?.ToString() == "Manual Anchor Plate Bar");
    }
}
