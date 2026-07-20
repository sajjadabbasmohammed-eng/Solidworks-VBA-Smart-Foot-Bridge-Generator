Attribute VB_Name = "modSmartBridgeTypesAndConstants"
Option Explicit

Public g_swApp As Object

Public Const SB_SW_DOC_PART As Long = 1
Public Const SB_SW_DOC_ASSEMBLY As Long = 2
Public Const SB_SW_DOC_DRAWING As Long = 3

Public Const SB_SW_DEFAULT_TEMPLATE_PART As Long = 1
Public Const SB_SW_DEFAULT_TEMPLATE_ASSEMBLY As Long = 2
Public Const SB_SW_DEFAULT_TEMPLATE_DRAWING As Long = 3

Public Const SB_SW_OPEN_SILENT As Long = 1
Public Const SB_SW_SAVE_AS_CURRENT_VERSION As Long = 0
Public Const SB_SW_SAVE_SILENT As Long = 1
Public Const SB_SW_ENDCOND_BLIND As Long = 0
Public Const SB_SW_CUSTOMINFO_TEXT As Long = 30
Public Const SB_SW_CUSTOMPROP_REPLACE As Long = 2
Public Const SB_SW_ADD_COMP_CURRENT_CONFIG As Long = 0
Public Const SB_SW_BOM_ANCHOR_TOPLEFT As Long = 1
Public Const SB_SW_BOM_TYPE_TOPLEVEL As Long = 1

Public Const SB_CROSS_MEMBER_X_M As Double = 0.08
Public Const SB_CROSS_MEMBER_Z_M As Double = 0.08
Public Const SB_DECK_THICKNESS_M As Double = 0.04
Public Const SB_POST_SIZE_M As Double = 0.06
Public Const SB_HANDRAIL_SIZE_M As Double = 0.05
Public Const SB_KICKPLATE_THICKNESS_M As Double = 0.012
Public Const SB_KICKPLATE_HEIGHT_M As Double = 0.15
Public Const SB_ENDPLATE_THICKNESS_M As Double = 0.012
Public Const SB_ENDPLATE_WIDTH_ALLOWANCE_M As Double = 0.2
Public Const SB_ENDPLATE_HEIGHT_ALLOWANCE_M As Double = 0.15
Public Const SB_ENDPLATE_HOLE_DIAMETER_M As Double = 0.014
Public Const SB_ENDPLATE_HOLE_EDGE_M As Double = 0.06

Public Type BridgeDesign
    ProjectName As String
    SaveFolder As String
    OutputFolder As String
    CreatedOn As String
    DataWorkbookPath As String

    SpanM As Double
    BridgeWidthM As Double
    MaterialName As String
    DeckName As String
    FinishName As String
    finishR As Long
    finishG As Long
    finishB As Long
    FinishMultiplier As Double

    BeamName As String
    BeamDepthM As Double
    BeamWidthM As Double
    BeamAreaM2 As Double

    CrossMemberCount As Long
    GuardPostCount As Long
    DeckPanelCount As Long
    ActualCrossSpacingM As Double
    RailHeightM As Double
    DeckPanelMaxLengthM As Double
    DeckPanelActualLengthM As Double

    EstimatedCostGBP As Double
    EstimatedMassKG As Double

    MainBeamPath As String
    CrossMemberPath As String
    DeckPanelPath As String
    GuardPostPath As String
    HandrailPath As String
    KickPlatePath As String
    EndPlatePath As String
    AssemblyPath As String
    drawingPath As String
    DrawingPdfPath As String
    DrawingDwgPath As String
    BomCsvPath As String
    ReportPath As String
End Type
