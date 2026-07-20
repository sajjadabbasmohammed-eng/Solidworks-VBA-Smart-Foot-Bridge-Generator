Attribute VB_Name = "modSmartBridgeAssembly"
Option Explicit

Public Sub CreateBridgeAssembly(ByRef d As BridgeDesign)

    Dim swModel As Object
    Dim swAssy As Object
    Dim beamY As Double
    Dim beamTopZ As Double
    Dim crossCenterZ As Double
    Dim deckCenterZ As Double
    Dim deckTopZ As Double
    Dim postCenterZ As Double
    Dim railTopZ As Double
    Dim railMidZ As Double
    Dim sideRailY As Double
    Dim kickPlateZ As Double
    Dim i As Long
    Dim side As Long
    Dim x As Double
    Dim nPostsPerSide As Long
    Dim postSpacingM As Double

    OpenModelSilent d.MainBeamPath, SB_SW_DOC_PART
    OpenModelSilent d.CrossMemberPath, SB_SW_DOC_PART
    OpenModelSilent d.DeckPanelPath, SB_SW_DOC_PART
    OpenModelSilent d.GuardPostPath, SB_SW_DOC_PART
    OpenModelSilent d.HandrailPath, SB_SW_DOC_PART
    OpenModelSilent d.KickPlatePath, SB_SW_DOC_PART
    OpenModelSilent d.EndPlatePath, SB_SW_DOC_PART

    Set swModel = NewModelDocument(SB_SW_DOC_ASSEMBLY)
    Set swAssy = swModel

    beamY = (d.BridgeWidthM / 2#) + (d.BeamWidthM / 2#)
    beamTopZ = d.BeamDepthM / 2#
    crossCenterZ = beamTopZ + (SB_CROSS_MEMBER_Z_M / 2#)
    deckCenterZ = beamTopZ + SB_CROSS_MEMBER_Z_M + (SB_DECK_THICKNESS_M / 2#)
    deckTopZ = beamTopZ + SB_CROSS_MEMBER_Z_M + SB_DECK_THICKNESS_M

    postCenterZ = deckTopZ + (d.RailHeightM / 2#)
    railTopZ = deckTopZ + d.RailHeightM
    railMidZ = deckTopZ + (d.RailHeightM * 0.55)
    kickPlateZ = deckTopZ + (SB_KICKPLATE_HEIGHT_M / 2#)


    sideRailY = (d.BridgeWidthM / 2#) - (SB_POST_SIZE_M / 2#)
    If sideRailY < SB_POST_SIZE_M Then sideRailY = d.BridgeWidthM / 2#

    InsertBridgeComponent swAssy, d.MainBeamPath, 0#, -beamY, 0#
    InsertBridgeComponent swAssy, d.MainBeamPath, 0#, beamY, 0#

    For i = 0 To d.CrossMemberCount - 1
        If d.CrossMemberCount = 1 Then
            x = 0#
        Else
            x = -(d.SpanM / 2#) + (CDbl(i) * d.ActualCrossSpacingM)
        End If
        InsertBridgeComponent swAssy, d.CrossMemberPath, x, 0#, crossCenterZ
    Next i

    For i = 0 To d.DeckPanelCount - 1
        x = -(d.SpanM / 2#) + (d.DeckPanelActualLengthM / 2#) + (CDbl(i) * d.DeckPanelActualLengthM)
        InsertBridgeComponent swAssy, d.DeckPanelPath, x, 0#, deckCenterZ
    Next i

    nPostsPerSide = d.GuardPostCount \ 2
    If nPostsPerSide < 2 Then nPostsPerSide = 2
    postSpacingM = d.SpanM / CDbl(nPostsPerSide - 1)

    For side = -1 To 1 Step 2
        For i = 0 To nPostsPerSide - 1
            x = -(d.SpanM / 2#) + (CDbl(i) * postSpacingM)
            InsertBridgeComponent swAssy, d.GuardPostPath, x, CDbl(side) * sideRailY, postCenterZ
        Next i
    Next side

    For side = -1 To 1 Step 2
        InsertBridgeComponent swAssy, d.HandrailPath, 0#, CDbl(side) * sideRailY, railTopZ
        InsertBridgeComponent swAssy, d.HandrailPath, 0#, CDbl(side) * sideRailY, railMidZ
        InsertBridgeComponent swAssy, d.KickPlatePath, 0#, CDbl(side) * sideRailY, kickPlateZ
    Next side

    InsertBridgeComponent swAssy, d.EndPlatePath, -(d.SpanM / 2#) - (SB_ENDPLATE_THICKNESS_M / 2#), 0#, 0#
    InsertBridgeComponent swAssy, d.EndPlatePath, (d.SpanM / 2#) + (SB_ENDPLATE_THICKNESS_M / 2#), 0#, 0#

    AddCustomPropertyText swModel, "Project", d.ProjectName
    AddCustomPropertyText swModel, "Description", "SmartBridge generated bridge assembly"
    AddCustomPropertyText swModel, "Span", FormatM(d.SpanM)
    AddCustomPropertyText swModel, "BridgeWidth", FormatM(d.BridgeWidthM)
    AddCustomPropertyText swModel, "Material", d.MaterialName
    AddCustomPropertyText swModel, "Finish", d.FinishName
    AddCustomPropertyText swModel, "FinishRGB", CStr(d.finishR) & "," & CStr(d.finishG) & "," & CStr(d.finishB)
    AddCustomPropertyText swModel, "SelectedBeam", d.BeamName
    AddCustomPropertyText swModel, "EstimatedCostGBP", Format$(d.EstimatedCostGBP, "0")
    AddCustomPropertyText swModel, "EstimatedMassKG", Format$(d.EstimatedMassKG, "0")

    ZoomAndRebuild swModel
    SaveModelAs swModel, d.AssemblyPath

End Sub

Private Function InsertBridgeComponent(ByVal swAssy As Object, ByVal filePath As String, ByVal x As Double, ByVal y As Double, ByVal z As Double) As Object

    Set InsertBridgeComponent = swAssy.AddComponent5(NormalizeFilePath(filePath), SB_SW_ADD_COMP_CURRENT_CONFIG, "", False, "", x, y, z)
    If InsertBridgeComponent Is Nothing Then Err.Raise vbObjectError + 1200, "InsertBridgeComponent", "Could not insert component:" & vbCrLf & filePath

End Function
