Attribute VB_Name = "modRibbon"
Option Explicit

' ==============================================================================
' modRibbon - CustomUI14 Ribbon Callbacks & Command Dispatcher (v2.1)
' ==============================================================================

Private mRibbon As IRibbonUI

Public Sub Ribbon_OnLoad(ByVal ribbon As IRibbonUI)
    Set mRibbon = ribbon
End Sub

Public Sub Ribbon_Invalidate()
    On Error Resume Next
    If Not mRibbon Is Nothing Then mRibbon.Invalidate
    On Error GoTo 0
End Sub

Public Sub Ribbon_OnAction(ByVal control As IRibbonControl)
    On Error GoTo ErrHandler
    Select Case control.ID
        ' --- Project & Spec ---
        Case "btnInitQS"
            modSettings.InitQSWorkspace
        Case "btnSpecSheet"
            modSpecSheet.CreateOrShowSpecSheet True
        Case "btnSettings"
            modSettings.CreateOrShowSettingsSheet True
        Case "btnPresetVN"
            modSettings.ApplyMeasurementPreset "TT13/2021/TT-BXD (VN)"
        Case "btnPresetPOMI"
            modSettings.ApplyMeasurementPreset "POMI / SMM7"
        Case "btnPresetCustom"
            modSettings.ApplyMeasurementPreset "Custom"
        Case "btnLangVI"
            modSettings.SetBookLanguage "VI"
        Case "btnLangEN"
            modSettings.SetBookLanguage "EN"
        Case "btnLangBI"
            modSettings.SetBookLanguage "BI"

        ' --- Wave 1: Single-Entry Structure, CAD Parser, BOQ Linker ---
        Case "btnQsData"
            modQsData.CreateOrShowQsDataSheet True
        Case "btnVisualRebar"
            modRebarConfig.OpenVisualRebarEditor True
        Case "btnRebarConfig"
            modRebarConfig.CreateOrShowRebarConfigSheet True
        Case "btnExpandRebar"
            modRebarConfig.ExpandRebarPresetsFromQsData True
        Case "btnInsertMainBar"
            modRebarConfig.InsertBarRowInBlock "MAIN", "D07", True
        Case "btnInsertStirBar"
            modRebarConfig.InsertBarRowInBlock "STIR", "D39", True
        Case "btnInsertCustomBar"
            modRebarConfig.InsertCustomBarRowOnConfig "D01", True
        Case "btnPickShapeForCfg", "btnInsertBar"
            modRebarTKT.ShowBarShapePicker
        Case "btnGenStruct"
            modQsGenerator.GenerateOrSyncAllStructure True
        Case "btnMarkManual"
            modQsGenerator.MarkSelectedRowAsManual
        Case "btnScanCadBeam"
            modCadParser.ScanCadBeamSectionToSheet
        Case "btnParseCadData"
            modCadParser.ParseRawCadDataToSheet True
        Case "btnBoqLinker"
            modBoqLinker.ShowBoqLinker
        Case "btnMasterBOQ"
            modTools.CreateMasterBOQSheet True

        ' --- Wave 2: Finishing & Infrastructure ---
        Case "btnCreateFinishInputs"
            modFinishingData.CreateFinishingInputSheets True
        Case "btnGenFinish"
            modFinishingData.GenerateFinishingDetail True
        Case "btnCreateInfraInputs"
            modInfraData.CreateOrShowInfraSheet True
        Case "btnFeedCadToInfra"
            modInfraData.FeedLengthsFromCadParsed True
        Case "btnGenInfra"
            modInfraData.GenerateInfraDetail True

        ' --- Material Lookup & Tools ---
        Case "btnSteelLookup", "btnQuickWeight"
            modCatalog.ShowSteelLookupForm
        Case "btnExportCatalog"
            modCatalog.ExportCatalogToActiveBook True
        Case "btnCleanFile"
            modTools.CleanWorkbookGarbage True
        Case "btnAuditBOQ"
            modTools.AuditWorkbookFormulas True
        Case "btnExportClean"
            modTools.ExportCleanXlsx
        Case "btnEngineTest"
            modEngineBridge.TestEnginePing True
        Case "btnAbout"
            modUnicode.MsgBoxU modUnicode.U("zOfit QS Pro v2.1 (Ki\u1ebfn tr\u00fac Lai: Excel Add-in + .NET 8 Engine)\n- Nh\u1eadp 1 l\u1ea7n QS_Data / QS_Spec -> Sinh *_Detail & TKT_* thu\u1ea7n c\u00f4ng th\u1ee9c\n- H\u1ed7 tr\u1ee3 T\u00e1ch CAD th\u00f4, Map-Link BOQ C\u0110T, Ho\u00e0n thi\u1ec7n & H\u1ea1 t\u1ea7ng\n- Xu\u1ea5t .xlsx s\u1ea1ch 100%, m\u1edf kh\u00f4ng c\u1ea7n Add-in."), vbInformation

        ' --- Rebar BBS (TKT Pro), Structural Steel & Cutting Optimizer ---
        Case "btnTKTMong"
            modRebarTKT.CreateTKTSheet "TKT_Mong", modUnicode.U("M\u00f3ng / \u0110\u00e0i c\u1ecdc"), True
        Case "btnTKTGiang"
            modRebarTKT.CreateTKTSheet "TKT_GiangMong", modUnicode.U("Gi\u1eb1ng m\u00f3ng / \u0110\u00e0 ki\u1ec1ng"), True
        Case "btnTKTCot"
            modRebarTKT.CreateTKTSheet "TKT_CotVach", modUnicode.U("C\u1ed9t / V\u00e1ch"), True
        Case "btnTKTDam"
            modRebarTKT.CreateTKTSheet "TKT_Dam", modUnicode.U("D\u1ea7m"), True
        Case "btnTKTSan"
            modRebarTKT.CreateTKTSheet "TKT_San", modUnicode.U("S\u00e0n / C\u1ea7u thang"), True
        Case "btnTKTKhac"
            modRebarTKT.CreateTKTSheet "TKT_Khac", modUnicode.U("C\u1ea5u ki\u1ec7n kh\u00e1c"), True
        Case "btnRowDup"
            modRebarTKT.DuplicateBarRow
        Case "btnRowUp"
            modRebarTKT.MoveBarRowUp
        Case "btnRowDown"
            modRebarTKT.MoveBarRowDown
        Case "btnRowDel"
            modRebarTKT.DeleteBarRow
        Case "btnRedrawShapes"
            modRebarTKT.RedrawActiveSheetShapes True
        Case "btnClearShapes"
            modRebarTKT.ClearActiveSheetShapes
        Case "btnCreateSteelShape"
            modSteelShape.CreateSteelShapeSheet True
        Case "btnCreateSteelPlate"
            modSteelShape.CreateSteelPlateSheet True
        Case "btnSummaryRebar"
            modRebarSummary.SummarizeAllSteel True
        Case "btnCutOptimize"
            modRebarSummary.OptimizeBarCutting1D True

        Case Else
            modUnicode.MsgBoxU "Unknown Ribbon control: " & control.ID, vbExclamation
    End Select
    Exit Sub
ErrHandler:
    modUnicode.MsgBoxU "Error [" & control.ID & "]: " & Err.Description, vbCritical
End Sub
