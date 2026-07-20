Attribute VB_Name = "modSmartBridgeParts"
Option Explicit

Public Sub CreateAllBridgeParts(ByRef d As BridgeDesign)

    Dim crossMemberY As Double
    Dim endPlateY As Double
    Dim endPlateZ As Double

    crossMemberY = d.BridgeWidthM + (2# * d.BeamWidthM) + 0.05
    endPlateY = d.BridgeWidthM + (2# * d.BeamWidthM) + SB_ENDPLATE_WIDTH_ALLOWANCE_M
    endPlateZ = d.BeamDepthM + SB_ENDPLATE_HEIGHT_ALLOWANCE_M

    CreateBoxPart d.MainBeamPath, d, "001", "Main beam", "Longitudinal main beam selected by structural validation: " & d.BeamName, d.SpanM, d.BeamWidthM, d.BeamDepthM, "2"
    CreateBoxPart d.CrossMemberPath, d, "002", "Cross member", "Automatic cross member spanning between the main beams", SB_CROSS_MEMBER_X_M, crossMemberY, SB_CROSS_MEMBER_Z_M, CStr(d.CrossMemberCount)
    CreateBoxPart d.DeckPanelPath, d, "003", "Deck panel", "Automatic deck panel generated from maximum panel length rule", d.DeckPanelActualLengthM, d.BridgeWidthM, SB_DECK_THICKNESS_M, CStr(d.DeckPanelCount)
    CreateBoxPart d.GuardPostPath, d, "004", "Guard post", "Vertical guard rail post generated from post spacing rule", SB_POST_SIZE_M, SB_POST_SIZE_M, d.RailHeightM, CStr(d.GuardPostCount)
    CreateBoxPart d.HandrailPath, d, "005", "Handrail", "Top and middle side rail generated automatically", d.SpanM, SB_HANDRAIL_SIZE_M, SB_HANDRAIL_SIZE_M, "4"
    CreateBoxPart d.KickPlatePath, d, "006", "Kick plate", "Side kick plate generated automatically", d.SpanM, SB_KICKPLATE_THICKNESS_M, SB_KICKPLATE_HEIGHT_M, "2"
    CreateBoxPart d.EndPlatePath, d, "007", "End plate", "Rectangular bridge end plate generated from selected beam and bridge width", SB_ENDPLATE_THICKNESS_M, endPlateY, endPlateZ, "2"

End Sub

Public Sub CreateBoxPart(ByVal filePath As String, ByRef d As BridgeDesign, ByVal itemNo As String, ByVal partTitle As String, ByVal description As String, ByVal sizeX As Double, ByVal sizeY As Double, ByVal sizeZ As Double, ByVal quantityText As String)

    Dim swModel As Object

    Set swModel = NewModelDocument(SB_SW_DOC_PART)
    CreateCenteredBoxFeature swModel, sizeX, sizeY, sizeZ
    ApplyAppearanceColor swModel, d.finishR, d.finishG, d.finishB
    TryApplySolidWorksMaterial swModel, d.MaterialName
    AddStandardPartProperties swModel, d, itemNo, partTitle, description, quantityText, sizeX, sizeY, sizeZ
    ZoomAndRebuild swModel
    SaveModelAs swModel, filePath

End Sub

Private Sub CreateCenteredBoxFeature(ByVal swModel As Object, ByVal sizeX As Double, ByVal sizeY As Double, ByVal sizeZ As Double)

    Dim ok As Boolean
    Dim swFeat As Object

    If sizeX <= 0# Or sizeY <= 0# Or sizeZ <= 0# Then Err.Raise vbObjectError + 1100, "CreateCenteredBoxFeature", "Part dimensions must be positive and non-zero."

    swModel.ClearSelection2 True
    ok = swModel.Extension.SelectByID2("Front Plane", "PLANE", 0#, 0#, 0#, False, 0, Nothing, 0)
    If Not ok Then Err.Raise vbObjectError + 1101, "CreateCenteredBoxFeature", "Could not select Front Plane."

    swModel.SketchManager.InsertSketch True
    swModel.SketchManager.CreateCenterRectangle 0#, 0#, 0#, sizeX / 2#, sizeY / 2#, 0#
    swModel.SketchManager.InsertSketch True

    Set swFeat = swModel.FeatureManager.FeatureExtrusion2(False, False, False, SB_SW_ENDCOND_BLIND, SB_SW_ENDCOND_BLIND, sizeZ / 2#, sizeZ / 2#, False, False, False, False, 0#, 0#, False, False, False, False, True, True, True, 0, 0#, False)
    If swFeat Is Nothing Then Err.Raise vbObjectError + 1102, "CreateCenteredBoxFeature", "SolidWorks failed to create the extrusion feature."

End Sub

Private Sub AddStandardPartProperties(ByVal swModel As Object, ByRef d As BridgeDesign, ByVal itemNo As String, ByVal partTitle As String, ByVal description As String, ByVal quantityText As String, ByVal sizeX As Double, ByVal sizeY As Double, ByVal sizeZ As Double)

    AddCustomPropertyText swModel, "Project", d.ProjectName
    AddCustomPropertyText swModel, "ItemNo", itemNo
    AddCustomPropertyText swModel, "PartTitle", partTitle
    AddCustomPropertyText swModel, "Description", description
    AddCustomPropertyText swModel, "Quantity", quantityText
    AddCustomPropertyText swModel, "Material", d.MaterialName
    AddCustomPropertyText swModel, "Finish", d.FinishName
    AddCustomPropertyText swModel, "FinishRGB", CStr(d.finishR) & "," & CStr(d.finishG) & "," & CStr(d.finishB)
    AddCustomPropertyText swModel, "FinishCostMultiplier", Format$(d.FinishMultiplier, "0.000")
    AddCustomPropertyText swModel, "SelectedBeam", d.BeamName
    AddCustomPropertyText swModel, "Envelope_X", FormatMM(sizeX)
    AddCustomPropertyText swModel, "Envelope_Y", FormatMM(sizeY)
    AddCustomPropertyText swModel, "Envelope_Z", FormatMM(sizeZ)
    AddCustomPropertyText swModel, "GeneratedBy", "SmartBridge API macro"

End Sub
