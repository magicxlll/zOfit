namespace ZOfit.Engine.Models;

using System.Text.Json.Serialization;

public sealed class IpcResponse<T>
{
    [JsonPropertyName("ok")]
    public bool Ok { get; set; }

    [JsonPropertyName("cmd")]
    public string Cmd { get; set; } = string.Empty;

    [JsonPropertyName("message")]
    public string Message { get; set; } = string.Empty;

    [JsonPropertyName("data")]
    public T? Data { get; set; }

    [JsonPropertyName("warnings")]
    public List<string> Warnings { get; set; } = new();
}

public sealed class PingRequest
{
    [JsonPropertyName("echo")]
    public string Echo { get; set; } = string.Empty;
}

public sealed class PingResponse
{
    [JsonPropertyName("version")]
    public string Version { get; set; } = "2.0.0";

    [JsonPropertyName("echo")]
    public string Echo { get; set; } = string.Empty;

    [JsonPropertyName("timestamp")]
    public string Timestamp { get; set; } = string.Empty;
}

public sealed class ProjectSpecDto
{
    // General & Elevations
    public string ProjectName { get; set; } = "zOfit Project";
    public string LangMode { get; set; } = "VI";
    public double ElevNatural { get; set; } = 19.70;
    public double ElevFfl { get; set; } = 20.00;
    public double TopsoilThkM { get; set; } = 0.10;
    public double ExcWorkSpaceFtgM { get; set; } = 0.70;
    public double ExcWorkSpaceBeamM { get; set; } = 0.50;
    public double ExcSlopeK { get; set; } = 1.30;
    public double ExcMachinePct { get; set; } = 0.90;
    public double ExcLaborPct { get; set; } = 0.10;
    public double BackfillCompK { get; set; } = 1.30;

    // Concrete & Formwork (Per-type Lean Concrete defaults + legacy aliases)
    public double LeanThkM { get; set; } = 0.10;
    public double LeanOverhangM { get; set; } = 0.10;
    public double LeanFtgThkM { get; set; } = 0.10;
    public double LeanFtgExtM { get; set; } = 0.10;
    public double LeanBeamThkM { get; set; } = 0.05;
    public double LeanBeamExtM { get; set; } = 0.10;
    public double LeanSlabThkM { get; set; } = 0.05;
    public double LeanSlabExtM { get; set; } = 0.00;
    public double LeanTankThkM { get; set; } = 0.10;
    public double LeanTankExtM { get; set; } = 0.15;
    public double ConcWasteFactor { get; set; } = 1.03;
    public double PeOverlapFactor { get; set; } = 1.01;
    public double OpeningDeductMinM2 { get; set; } = 0.50;
    public string BeamSlabRule { get; set; } = "1 - FloorFull_BeamMinusSlab";
    public string ColBeamRule { get; set; } = "1 - ColThrough_BeamMinusCol";
    public string FormDeduction { get; set; } = "Yes";

    // Rebar Standards
    public double CoverFtgMm { get; set; } = 50;
    public double CoverColMm { get; set; } = 35;
    public double CoverBeamMm { get; set; } = 30;
    public double CoverSlabMm { get; set; } = 20;
    public double CoverWallMm { get; set; } = 30;
    public double LanTensionMult { get; set; } = 40;
    public double LanCompMult { get; set; } = 30;
    public double LlapMult { get; set; } = 45;
    public double ColStaggerOffsetMm { get; set; } = 600;
    public double CouplerMinD { get; set; } = 20;
    public double StdBarLenMm { get; set; } = 11700;
    public double BendDeduct90Mult { get; set; } = 1;
    public double ScrapSmallPct { get; set; } = 3.0;   // <= D10
    public double ScrapMedPct { get; set; } = 2.5;     // D12-D18
    public double ScrapLargePct { get; set; } = 2.0;   // > D18

    // Pile & Industrial Slab
    public double PileCutHeadM { get; set; } = 0.50;
    public double PileSegmentLenM { get; set; } = 11.0;
    public double JointGridSpacingM { get; set; } = 6.0;

    // Finishing & Infra (Wave 2)
    public double WallOpeningMinM2 { get; set; } = 0.25;
    public double CrackMeshWidthM { get; set; } = 0.20;
    public double WaterproofCurbHM { get; set; } = 0.30;
    public double PipeBeddingM { get; set; } = 0.15;
    public double PipeSlopeM { get; set; } = 0.30;
    public double RoadDowelLenM { get; set; } = 0.50;
    public double RoadDowelSpacingM { get; set; } = 0.30;
}

public sealed class ElementDataDto
{
    public int DataRow { get; set; }
    public string ElementId { get; set; } = string.Empty;
    public string Building { get; set; } = string.Empty;
    public string Group { get; set; } = string.Empty;
    public string Mark { get; set; } = string.Empty;
    public string Description { get; set; } = string.Empty;
    public double Qty { get; set; }
    public double LengthM { get; set; }
    public double WidthM { get; set; }
    public double HeightM { get; set; }
    public double TopElevM { get; set; }
    public double BotElevM { get; set; }
    public string ConcGrade { get; set; } = "M400";
    public double ClashDeductVolM3 { get; set; }
    public double ClashDeductFwM2 { get; set; }
    public string IsBuried { get; set; } = "N";
    public double? LeanThkOvrM { get; set; }
    public double? LeanExtOvrM { get; set; }
    public string RebarTemplate { get; set; } = string.Empty; // Preset name on QS_Data Col R
    // Optional inline quick-seed fields for fallback/preset generation
    public double MainXDia { get; set; }
    public string MainXSpec { get; set; } = string.Empty;
    public double MainYDia { get; set; }
    public string MainYSpec { get; set; } = string.Empty;
    public double TopXDia { get; set; }
    public string TopXSpec { get; set; } = string.Empty;
    public double TopYDia { get; set; }
    public string TopYSpec { get; set; } = string.Empty;
    public double StirrupDia { get; set; }
    public string StirrupSpec { get; set; } = string.Empty;
    public string ExtraSpec { get; set; } = string.Empty;
    public string BoqBucket { get; set; } = "Substructure";
}

public sealed class RebarConfigRowDto
{
    public int ConfigRow { get; set; } // 1-based row on QS_RebarConfig
    public string RowKind { get; set; } = "BAR"; // BAR | BLK_HDR | BLK_END
    public string ElementId { get; set; } = string.Empty;
    public string Building { get; set; } = string.Empty;
    public string Group { get; set; } = string.Empty;
    public string BarId { get; set; } = string.Empty;
    public string Role { get; set; } = "CUSTOM";
    public string ShapeCode { get; set; } = "D01";
    public double Dia { get; set; } = 14;
    public double LayerMult { get; set; } = 1; // Part / Mặt

    // Transparent Parametric BBS v3 fields (Zone B)
    public string BarDir { get; set; } = string.Empty;       // X | Y | H | X*0.25 | Y*0.25 | X*0.7 | PERIM_YH | PERIM_XY | PERIM_XY_INNER | TIE_Y | TIE_X | SLOPED_XH | CHAIR_H | CUSTOM
    public string DistDir { get; set; } = string.Empty;      // X | Y | H | ZONE2_X | ZONE2_H | XY_GRID | XH_GRID | FIXED
    public string HookLExpr { get; set; } = string.Empty;    // <50 => xD, >=50 => mm
    public string CoverExpr { get; set; } = string.Empty;    // =SPEC_Cover_* or numeric mm
    public string HookRExpr { get; set; } = string.Empty;    // <50 => xD, >=50 => mm
    public double? Space1Mm { get; set; }
    public double? Space2Mm { get; set; }
    public string OverrideQtyExpr { get; set; } = string.Empty;  // SL.m
    public string OverrideLenExpr { get; set; } = string.Empty;  // L.m
    public string DeltaDistExpr { get; set; } = string.Empty;    // ±Pb (mm)
    public string DeltaLenExpr { get; set; } = string.Empty;     // ±Lth (mm)
    public double? Le1aMm { get; set; }
    public double? Le1bMm { get; set; }
    public double? Ri2bMm { get; set; }
    public double? Ri2aMm { get; set; }

    // Evaluated Zone C preview values (optional when read from Excel, or legacy v2 fields)
    public string CountOrSpace { get; set; } = "4";
    public string DistMode { get; set; } = "AUTO";
    public string LenMode { get; set; } = "AUTO";
    public double CustomL1Mm { get; set; }
    public double CustomL2Mm { get; set; }
    public double CustomL3Mm { get; set; }
    public double StaggerOrExtraMm { get; set; }
    public string Note { get; set; } = string.Empty;
}

public sealed class ExpandRebarPresetsRequest
{
    public ProjectSpecDto Spec { get; set; } = new();
    public List<ElementDataDto> Elements { get; set; } = new();
    public List<RebarConfigRowDto> ExistingConfigs { get; set; } = new();
}

public sealed class ExpandRebarPresetsResponse
{
    public int AddedRowsCount { get; set; }
    public List<RebarConfigRowDto> Rows { get; set; } = new();
}

public sealed class ExistingDetailRowDto
{
    public int RowIndex { get; set; }
    public string Code { get; set; } = string.Empty;
    public string Description { get; set; } = string.Empty;
    public string SpecText { get; set; } = string.Empty;
    public string Unit { get; set; } = string.Empty;
    public string NoExpr { get; set; } = string.Empty;
    public string LExpr { get; set; } = string.Empty;
    public string WExpr { get; set; } = string.Empty;
    public string HExpr { get; set; } = string.Empty;
    public string AExpr { get; set; } = string.Empty;
    public string BExpr { get; set; } = string.Empty;
    public string CExpr { get; set; } = string.Empty;
    public string QtyExpr { get; set; } = string.Empty;
    public string TagId { get; set; } = string.Empty; // ElementId | SECTION | SUBITEM | MANUAL
    public string SectionKey { get; set; } = string.Empty;
}

public sealed class ExistingTktRowDto
{
    public int RowIndex { get; set; }
    public string CompName { get; set; } = string.Empty;
    public string BarMark { get; set; } = string.Empty;
    public string ShapeCode { get; set; } = string.Empty;
    public string L1Expr { get; set; } = string.Empty;
    public string L2Expr { get; set; } = string.Empty;
    public string L3Expr { get; set; } = string.Empty;
    public string L4Expr { get; set; } = string.Empty;
    public string L5Expr { get; set; } = string.Empty;
    public string L6Expr { get; set; } = string.Empty;
    public string L7Expr { get; set; } = string.Empty;
    public string M1Expr { get; set; } = string.Empty;
    public string M2Expr { get; set; } = string.Empty;
    public double Dia { get; set; }
    public string QtyPerElExpr { get; set; } = string.Empty;
    public string ElQtyExpr { get; set; } = string.Empty;
    public string Note { get; set; } = string.Empty;
    public string TagId { get; set; } = string.Empty; // ElementId::BarKey or MANUAL
    public string BoqBucket { get; set; } = string.Empty;
}

public sealed class GenerateStructureRequest
{
    public string TargetBuilding { get; set; } = string.Empty; // empty => generate first or all
    public ProjectSpecDto Spec { get; set; } = new();
    public List<ElementDataDto> Elements { get; set; } = new();
    public List<RebarConfigRowDto> RebarConfigs { get; set; } = new();
    public List<ExistingDetailRowDto> ExistingDetailRows { get; set; } = new();
    public List<ExistingTktRowDto> ExistingTktRows { get; set; } = new();
}

public sealed class SheetCellRowDto
{
    // RowKind: TITLE | BANNER | WORK_HDR | SUB_HDR | ITEM_HDR | DATA | BLANK | MANUAL | REMOVED
    public string RowKind { get; set; } = "DATA";
    public int OutlineLevel { get; set; } = 0; // 0, 1, 2
    public string TagId { get; set; } = string.Empty;
    public string SectionKey { get; set; } = string.Empty;
    public List<object?> Values { get; set; } = new(); // Strings starting with "=" are Excel formulas!
}

public sealed class GeneratedSheetDto
{
    public string SheetName { get; set; } = string.Empty;
    public string Building { get; set; } = string.Empty;
    public int AddedCount { get; set; }
    public int UpdatedCount { get; set; }
    public int PreservedManualCount { get; set; }
    public int RemovedFlaggedCount { get; set; }
    public List<SheetCellRowDto> Rows { get; set; } = new();
}

public sealed class GenerateStructureResponse
{
    public List<GeneratedSheetDto> DetailSheets { get; set; } = new();
    public List<GeneratedSheetDto> TktSheets { get; set; } = new();
}

// CAD Parser Models
public sealed class CadRawRowDto
{
    public int SourceRow { get; set; }
    public string ColA { get; set; } = string.Empty; // Optional Block Name or Text
    public string ColB { get; set; } = string.Empty; // Optional MG
    public string ColC { get; set; } = string.Empty; // Optional DC
    public string ColD { get; set; } = string.Empty; // Optional Qty
}

public sealed class CadParseRequest
{
    public List<CadRawRowDto> RawRows { get; set; } = new();
}

public sealed class CadParsedItemDto
{
    public int SourceRow { get; set; }
    public string RawText { get; set; } = string.Empty;
    public string ItemType { get; set; } = string.Empty; // TRENCH | BOX_CULVERT | PIPE | MANHOLE | UNKNOWN
    public string CategoryKey { get; set; } = string.Empty; // e.g. A-W300, Box:600x600, D200, TN-TENGA2
    public double WidthMm { get; set; }
    public double HeightMm { get; set; }
    public double DiameterMm { get; set; }
    public double SlopePct { get; set; }
    public double LengthM { get; set; }
    public double MgElev { get; set; }
    public double DcElev { get; set; }
    public double DepthM { get; set; }
    public double Qty { get; set; } = 1;
    public bool NeedsManualCheck { get; set; }
}

public sealed class CadParseResponse
{
    public int TotalInput { get; set; }
    public int ParsedCount { get; set; }
    public int UnknownCount { get; set; }
    public List<CadParsedItemDto> Items { get; set; } = new();
    public List<string> DistinctKeys { get; set; } = new();
}

// Wave 2: Finishing Models
public sealed class DoorWinDto
{
    public int DataRow { get; set; }
    public string Code { get; set; } = string.Empty;
    public string DoorType { get; set; } = "DOOR";
    public double WidthM { get; set; }
    public double HeightM { get; set; }
    public double LintelExtendM { get; set; } = 0.25;
}

public sealed class RoomFinishDto
{
    public int DataRow { get; set; }
    public string RoomId { get; set; } = string.Empty;
    public string Building { get; set; } = string.Empty;
    public string Floor { get; set; } = string.Empty;
    public string RoomName { get; set; } = string.Empty;
    public double AreaM2 { get; set; }
    public double PerimM { get; set; }
    public double HeightM { get; set; }
    public string WallType { get; set; } = "B200"; // B100 | B200 | PANEL | NONE
    public double WallLenM { get; set; }
    public string DoorCodes { get; set; } = string.Empty; // comma separated: 1SD1*2,WD1*3
    public string FloorFinish { get; set; } = "TILE"; // TILE | EPOXY | HARDENER | SCREED
    public double WallTileHeightM { get; set; } = 0;
    public string CeilingType { get; set; } = "GYPSUM"; // GYPSUM | PAINT | NONE
    public string IsWetArea { get; set; } = "N";
    public double RoofAreaM2 { get; set; } = 0;
    public double ParapetLenM { get; set; } = 0;
    public double ParapetHeightM { get; set; } = 1.0;
}

public sealed class GenerateFinishingRequest
{
    public ProjectSpecDto Spec { get; set; } = new();
    public List<DoorWinDto> Doors { get; set; } = new();
    public List<RoomFinishDto> Rooms { get; set; } = new();
    public List<ExistingDetailRowDto> ExistingRows { get; set; } = new();
}

public sealed class GenerateFinishingResponse
{
    public GeneratedSheetDto Sheet { get; set; } = new();
}

// Wave 2: Infrastructure Models
public sealed class InfraItemDto
{
    public int DataRow { get; set; }
    public string InfraId { get; set; } = string.Empty;
    public string Category { get; set; } = string.Empty; // PIPE | CULVERT | ROAD_BTXM | ROAD_ASPHALT | FENCE
    public string Code { get; set; } = string.Empty;
    public string Description { get; set; } = string.Empty;
    public double Qty { get; set; } = 1;
    public double LengthM { get; set; }
    public double WidthM { get; set; }
    public double DepthOrThkM { get; set; }
    public double DiaMm { get; set; }
    public double SubBaseT1M { get; set; } = 0.15;
    public double SubBaseT2M { get; set; } = 0.15;
    public double CushionThkM { get; set; } = 0.15;
    public double ExtraSidewalkWM { get; set; } = 0;
    public double CurbLenM { get; set; } = 0;
    public double ParamA { get; set; } = 0; // joint spacing or fence brick H
    public double ParamB { get; set; } = 0; // col spacing or stone/top thk
}

public sealed class GenerateInfraRequest
{
    public ProjectSpecDto Spec { get; set; } = new();
    public List<InfraItemDto> Items { get; set; } = new();
    public List<ExistingDetailRowDto> ExistingRows { get; set; } = new();
}

public sealed class GenerateInfraResponse
{
    public GeneratedSheetDto Sheet { get; set; } = new();
}

public sealed class EditRebarUiRequest
{
    public long ParentHwnd { get; set; }
    public ProjectSpecDto Spec { get; set; } = new();
    public ElementDataDto Element { get; set; } = new();
    public List<RebarConfigRowDto> ExistingRows { get; set; } = new();
}

public sealed class EditRebarUiResponse
{
    public bool Accepted { get; set; }
    public string ElementId { get; set; } = string.Empty;
    public List<RebarConfigRowDto> Rows { get; set; } = new();
}

// Wave 3: CAD Smart Scan & Beam Section Models
public sealed class CadEntityItemDto
{
    public string Handle { get; set; } = string.Empty;
    public string Type { get; set; } = string.Empty; // TEXT | MTEXT | INSERT | DIMENSION | LWPOLYLINE
    public string Layer { get; set; } = string.Empty;
    public string Text { get; set; } = string.Empty;
    public double[] Position { get; set; } = Array.Empty<double>(); // [x, y, z]
    public CadBoundsDto? Bounds { get; set; }
    public string BlockName { get; set; } = string.Empty;
    public Dictionary<string, string>? Attributes { get; set; }
}

public sealed class CadBoundsDto
{
    public double[] Min { get; set; } = Array.Empty<double>();
    public double[] Max { get; set; } = Array.Empty<double>();
    public double Width { get; set; }
    public double Height { get; set; }
}

public sealed class CadBeamSectionParseRequest
{
    public string Building { get; set; } = "MAIN";
    public string Floor { get; set; } = "TANG1";
    public double DefaultLengthM { get; set; } = 6.0;
    public List<CadEntityItemDto> Entities { get; set; } = new();
    public List<string> RawTextLines { get; set; } = new();
}

public sealed class BeamSectionDetailsDto
{
    public string BeamName { get; set; } = string.Empty;
    public double WidthMm { get; set; } = 800;
    public double HeightMm { get; set; } = 800;
    public double TopElevM { get; set; } = 0.0;
    public string TopBarLeft { get; set; } = string.Empty;
    public string TopBarMid { get; set; } = string.Empty;
    public string TopBarRight { get; set; } = string.Empty;
    public string BotBarLeft { get; set; } = string.Empty;
    public string BotBarMid { get; set; } = string.Empty;
    public string BotBarRight { get; set; } = string.Empty;
    public string StirrupLeft { get; set; } = string.Empty;
    public string StirrupMid { get; set; } = string.Empty;
    public string StirrupRight { get; set; } = string.Empty;
    public string SideBar { get; set; } = string.Empty;
    public string SupportBarLeft { get; set; } = string.Empty;
    public string SupportBarMid { get; set; } = string.Empty;
    public string SupportBarRight { get; set; } = string.Empty;
}

public sealed class ParsedBeamResultDto
{
    public ElementDataDto Element { get; set; } = new();
    public List<RebarConfigRowDto> RebarRows { get; set; } = new();
    public BeamSectionDetailsDto SectionDetails { get; set; } = new();
}

public sealed class CadBeamSectionParseResponse
{
    public int BeamCount { get; set; }
    public List<ParsedBeamResultDto> Beams { get; set; } = new();
    public List<string> Messages { get; set; } = new();
}

public sealed class ScanCadBeamRequest
{
    public long ParentHwnd { get; set; }
    public string Host { get; set; } = "127.0.0.1";
    public int Port { get; set; } = 9877;
    public string Building { get; set; } = "MAIN";
    public string Floor { get; set; } = "TANG1";
    public double DefaultLengthM { get; set; } = 6.0;
    public bool AutoOpenPreview { get; set; } = true;
    public ProjectSpecDto Spec { get; set; } = new();
}

public sealed class ScanCadBeamResponse
{
    public bool Accepted { get; set; }
    public int BeamCount { get; set; }
    public List<ParsedBeamResultDto> Beams { get; set; } = new();
    public string Message { get; set; } = string.Empty;
}

