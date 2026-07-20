Attribute VB_Name = "modSmartBridgeMain"
Option Explicit

Private mBuildStage As String

Public Sub main()
    frmSmartBridge.Show
End Sub

Public Sub BuildBridgeFromForm(ByVal f As Object)

    Dim d As BridgeDesign
    Dim errNo As Long
    Dim errSrc As String
    Dim errDesc As String
    Dim msg As String

    On Error GoTo BuildFail

    mBuildStage = "checking UserForm validation"
    If f Is Nothing Then Err.Raise vbObjectError + 1300, "BuildBridgeFromForm", "No UserForm object was supplied."
    If Not f.IsStructurallySound Then Err.Raise vbObjectError + 1301, "BuildBridgeFromForm", "The bridge has not passed the structural validation box."

    mBuildStage = "connecting to SolidWorks"
    If Not ConnectToSolidWorks() Then Exit Sub

    SafeStartCommand

    mBuildStage = "reading validated values from UserForm"
    FillDesignFromForm f, d

    mBuildStage = "preparing output folder"
    PrepareDesignPaths d

    mBuildStage = "creating generated part files"
    CreateAllBridgeParts d

    mBuildStage = "creating bridge assembly"
    CreateBridgeAssembly d

    mBuildStage = "creating bridge drawing"
    CreateBridgeDrawing d
    
    mBuildStage = "creating dimensioned part drawings and DWG files"
    CreatePartDrawings d
    
    mBuildStage = "writing CSV BOM"
    WriteBridgeBomCsv d

    mBuildStage = "writing text report"
    WriteBridgeReportTxt d

    SafeEndCommand

    msg = "SmartBridge model generation complete." & vbCrLf & vbCrLf
    msg = msg & "Output folder:" & vbCrLf & d.OutputFolder & vbCrLf & vbCrLf
    msg = msg & "Created: parts, assembly, drawing, PDF/DWG if supported, CSV BOM and text report."
    MsgBox msg, vbInformation, "SmartBridge API"

    Exit Sub

BuildFail:

    errNo = Err.Number
    errSrc = Err.Source
    errDesc = Err.description

    SafeEndCommand

    msg = "SmartBridge model generation failed." & vbCrLf & vbCrLf
    msg = msg & "Failed stage:" & vbCrLf & mBuildStage & vbCrLf & vbCrLf
    msg = msg & "Error number: " & CStr(errNo) & vbCrLf
    msg = msg & "Error source: " & errSrc & vbCrLf
    msg = msg & "Details: " & errDesc
    MsgBox msg, vbCritical, "SmartBridge API"

End Sub

Private Sub SafeStartCommand()
    On Error Resume Next
    If Not g_swApp Is Nothing Then g_swApp.CommandInProgress = True
    On Error GoTo 0
End Sub

Private Sub SafeEndCommand()
    On Error Resume Next
    If Not g_swApp Is Nothing Then g_swApp.CommandInProgress = False
    On Error GoTo 0
End Sub

Private Sub FillDesignFromForm(ByVal f As Object, ByRef d As BridgeDesign)

    d.ProjectName = CleanFileName(CStr(f.ProjectName))
    d.SaveFolder = CStr(f.SaveFolder)
    d.DataWorkbookPath = CStr(f.DataWorkbookPath)

    If Len(Trim$(d.ProjectName)) = 0 Then d.ProjectName = "SmartBridge"
    If Len(Trim$(d.SaveFolder)) = 0 Then d.SaveFolder = Environ$("USERPROFILE") & "\Documents\SmartBridgeAPI\Output"

    d.SpanM = CDbl(f.SpanM)
    d.BridgeWidthM = CDbl(f.BridgeWidthM)
    d.MaterialName = CStr(f.MaterialName)
    d.DeckName = CStr(f.DeckName)
    d.FinishName = CStr(f.FinishName)
    d.finishR = CLng(f.FinishRed)
    d.finishG = CLng(f.FinishGreen)
    d.finishB = CLng(f.FinishBlue)
    d.FinishMultiplier = CDbl(f.FinishMultiplier)

    d.BeamName = CStr(f.SelectedBeamName)
    d.BeamDepthM = CDbl(f.SelectedBeamDepthM)
    d.BeamWidthM = CDbl(f.SelectedBeamWidthM)
    d.BeamAreaM2 = CDbl(f.SelectedBeamAreaM2)

    d.CrossMemberCount = CLng(f.CrossMemberCount)
    d.GuardPostCount = CLng(f.GuardPostCount)
    d.DeckPanelCount = CLng(f.DeckPanelCount)
    d.ActualCrossSpacingM = CDbl(f.ActualCrossSpacingM)
    d.RailHeightM = CDbl(f.RailHeightM)
    d.DeckPanelMaxLengthM = CDbl(f.DeckPanelMaxLengthM)
    d.EstimatedCostGBP = CDbl(f.EstimatedCostGBP)
    d.EstimatedMassKG = CDbl(f.EstimatedMassKG)

    If d.SpanM <= 0# Then Err.Raise vbObjectError + 1310, "FillDesignFromForm", "Span is zero or negative."
    If d.BridgeWidthM <= 0# Then Err.Raise vbObjectError + 1311, "FillDesignFromForm", "Bridge width is zero or negative."
    If d.BeamDepthM <= 0# Then Err.Raise vbObjectError + 1312, "FillDesignFromForm", "Selected beam depth is invalid."
    If d.BeamWidthM <= 0# Then Err.Raise vbObjectError + 1313, "FillDesignFromForm", "Selected beam width is invalid."
    If d.BeamAreaM2 <= 0# Then Err.Raise vbObjectError + 1314, "FillDesignFromForm", "Selected beam area is invalid."

    If d.CrossMemberCount < 2 Then d.CrossMemberCount = 2
    If d.GuardPostCount < 4 Then d.GuardPostCount = 4
    If d.DeckPanelCount < 1 Then d.DeckPanelCount = 1
    If d.RailHeightM <= 0# Then d.RailHeightM = 1.1
    If d.FinishMultiplier <= 0# Then d.FinishMultiplier = 1#

    d.DeckPanelActualLengthM = d.SpanM / CDbl(d.DeckPanelCount)

End Sub
