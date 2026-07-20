Attribute VB_Name = "modSmartBridgeDrawing"
Option Explicit

Public Sub CreateBridgeDrawing(ByRef d As BridgeDesign)

    Dim swModel As Object
    Dim swDraw As Object
    Dim swSheet As Object
    Dim props As Variant

    Dim sheetW As Double
    Dim sheetH As Double
    Dim viewScale As Double

    Dim vIso As Object
    Dim vFront As Object
    Dim vTop As Object
    Dim vRight As Object

    OpenModelSilent d.AssemblyPath, SB_SW_DOC_ASSEMBLY

    Set swModel = NewModelDocument(SB_SW_DOC_DRAWING)
    Set swDraw = swModel

    sheetW = 0.42
    sheetH = 0.297

    On Error Resume Next
    Set swSheet = swDraw.GetCurrentSheet
    props = swSheet.GetProperties

    If IsArray(props) Then
        If UBound(props) >= 6 Then
            If CDbl(props(5)) > 0# Then sheetW = CDbl(props(5))
            If CDbl(props(6)) > 0# Then sheetH = CDbl(props(6))
        End If
    End If
    On Error GoTo 0

    viewScale = CalculateSafeDrawingScale(d, sheetW, sheetH)

    
    Set vFront = swDraw.CreateDrawViewFromModelView3(d.AssemblyPath, "*Front", sheetW * 0.31, sheetH * 0.64, 0#)
    Set vTop = swDraw.CreateDrawViewFromModelView3(d.AssemblyPath, "*Top", sheetW * 0.31, sheetH * 0.32, 0#)
    Set vIso = swDraw.CreateDrawViewFromModelView3(d.AssemblyPath, "*Isometric", sheetW * 0.73, sheetH * 0.66, 0#)
    Set vRight = swDraw.CreateDrawViewFromModelView3(d.AssemblyPath, "*Right", sheetW * 0.73, sheetH * 0.34, 0#)

    ApplyViewScale vFront, viewScale
    ApplyViewScale vTop, viewScale
    ApplyViewScale vIso, viewScale * 0.85
    ApplyViewScale vRight, viewScale

    
    AddDrawingDimensionPack swModel, d, sheetW, sheetH

    If Not vIso Is Nothing Then
        TryInsertBomTable vIso, sheetW, sheetH
    End If

    AddCustomPropertyText swModel, "Project", d.ProjectName
    AddCustomPropertyText swModel, "Description", "SmartBridge automatic dimensioned drawing"
    AddCustomPropertyText swModel, "Span", FormatM(d.SpanM)
    AddCustomPropertyText swModel, "BridgeWidth", FormatM(d.BridgeWidthM)
    AddCustomPropertyText swModel, "RailHeight", FormatM(d.RailHeightM)
    AddCustomPropertyText swModel, "Material", d.MaterialName
    AddCustomPropertyText swModel, "Finish", d.FinishName
    AddCustomPropertyText swModel, "SelectedBeam", d.BeamName

    ZoomAndRebuild swModel

    SaveModelAs swModel, d.drawingPath

    On Error Resume Next
    SaveModelAs swModel, d.DrawingPdfPath
    SaveModelAs swModel, d.DrawingDwgPath
    On Error GoTo 0

End Sub

Private Function CalculateSafeDrawingScale(ByRef d As BridgeDesign, ByVal sheetW As Double, ByVal sheetH As Double) As Double

    Dim availableW As Double
    Dim availableH As Double
    Dim modelW As Double
    Dim modelH As Double
    Dim s1 As Double
    Dim s2 As Double
    Dim s As Double

    availableW = sheetW * 0.36
    availableH = sheetH * 0.22

    modelW = d.SpanM
    modelH = d.RailHeightM + d.BeamDepthM + 0.35

    If modelW <= 0# Then modelW = 1#
    If modelH <= 0# Then modelH = 1#

    s1 = availableW / modelW
    s2 = availableH / modelH

    If s1 < s2 Then
        s = s1
    Else
        s = s2
    End If

    If s > 0.1 Then s = 0.1
    If s < 0.003 Then s = 0.003

    CalculateSafeDrawingScale = s

End Function

Private Sub ApplyViewScale(ByVal swView As Object, ByVal scaleValue As Double)

    On Error Resume Next

    If Not swView Is Nothing Then
        swView.UseSheetScale = False
        swView.ScaleDecimal = scaleValue
    End If

    On Error GoTo 0

End Sub

Private Sub TryInsertBomTable(ByVal swView As Object, ByVal sheetW As Double, ByVal sheetH As Double)

    Dim bomAnn As Object

    On Error Resume Next

    Set bomAnn = swView.InsertBomTable2(True, sheetW * 0.04, sheetH * 0.94, SB_SW_BOM_ANCHOR_TOPLEFT, SB_SW_BOM_TYPE_TOPLEVEL, "", "")

    On Error GoTo 0

End Sub



Private Sub AddDrawingDimensionPack(ByVal swModel As Object, ByRef d As BridgeDesign, ByVal sheetW As Double, ByVal sheetH As Double)

    Dim xLeft As Double
    Dim xRight As Double
    Dim yTop As Double
    Dim yBottom As Double
    Dim scheduleX As Double
    Dim scheduleY As Double
    Dim textSize As Double

    textSize = 0.003

    xLeft = sheetW * 0.08
    xRight = sheetW * 0.55
    yTop = sheetH * 0.88
    yBottom = sheetH * 0.08

    scheduleX = sheetW * 0.56
    scheduleY = sheetH * 0.22

    
    AddDrawingNote swModel, "SMARTBRIDGE DIMENSIONED DRAWING", sheetW * 0.04, sheetH * 0.975, textSize * 1.15

    
    AddDrawingNote swModel, "OVERALL SPAN = " & FormatM(d.SpanM), sheetW * 0.15, sheetH * 0.5, textSize
    AddDrawingNote swModel, "BRIDGE WIDTH = " & FormatM(d.BridgeWidthM), sheetW * 0.15, sheetH * 0.18, textSize
    AddDrawingNote swModel, "RAIL HEIGHT = " & FormatM(d.RailHeightM), sheetW * 0.67, sheetH * 0.49, textSize
    AddDrawingNote swModel, "BEAM DEPTH = " & FormatMM(d.BeamDepthM), sheetW * 0.67, sheetH * 0.2, textSize
    AddDrawingNote swModel, "BEAM WIDTH = " & FormatMM(d.BeamWidthM), sheetW * 0.67, sheetH * 0.16, textSize

    
    AddDrawingLine swModel, sheetW * 0.1, sheetH * 0.47, sheetW * 0.5, sheetH * 0.47
    AddDrawingLine swModel, sheetW * 0.1, sheetH * 0.15, sheetW * 0.5, sheetH * 0.15
    AddDrawingLine swModel, sheetW * 0.64, sheetH * 0.26, sheetW * 0.64, sheetH * 0.46

    AddEndTicks swModel, sheetW * 0.1, sheetH * 0.47, sheetW * 0.5, sheetH * 0.47
    AddEndTicks swModel, sheetW * 0.1, sheetH * 0.15, sheetW * 0.5, sheetH * 0.15
    AddEndTicks swModel, sheetW * 0.64, sheetH * 0.26, sheetW * 0.64, sheetH * 0.46

    
    AddDimensionSchedule swModel, d, scheduleX, scheduleY, textSize

End Sub

Private Sub AddDimensionSchedule(ByVal swModel As Object, ByRef d As BridgeDesign, ByVal x As Double, ByVal y As Double, ByVal textSize As Double)

    Dim s As String

    s = "DIMENSION SCHEDULE" & vbCrLf
    s = s & "--------------------------" & vbCrLf
    s = s & "Span: " & FormatM(d.SpanM) & vbCrLf
    s = s & "Bridge width: " & FormatM(d.BridgeWidthM) & vbCrLf
    s = s & "Rail height: " & FormatM(d.RailHeightM) & vbCrLf
    s = s & "Selected beam: " & d.BeamName & vbCrLf
    s = s & "Beam depth: " & FormatMM(d.BeamDepthM) & vbCrLf
    s = s & "Beam width: " & FormatMM(d.BeamWidthM) & vbCrLf
    s = s & "Cross members: " & CStr(d.CrossMemberCount) & vbCrLf
    s = s & "Cross spacing: " & FormatMM(d.ActualCrossSpacingM) & vbCrLf
    s = s & "Deck panels: " & CStr(d.DeckPanelCount) & vbCrLf
    s = s & "Deck panel length: " & FormatM(d.DeckPanelActualLengthM) & vbCrLf
    s = s & "Guard posts: " & CStr(d.GuardPostCount) & vbCrLf
    s = s & "Material: " & d.MaterialName & vbCrLf
    s = s & "Finish: " & d.FinishName

    AddDrawingNote swModel, s, x, y, textSize

End Sub

Private Sub AddDrawingNote(ByVal swModel As Object, ByVal noteText As String, ByVal x As Double, ByVal y As Double, ByVal textSize As Double)

    Dim swNote As Object
    Dim swAnn As Object
    Dim swTextFormat As Object

    On Error Resume Next

    Set swNote = swModel.InsertNote(noteText)

    If Not swNote Is Nothing Then

        Set swAnn = swNote.GetAnnotation

        If Not swAnn Is Nothing Then
            swAnn.SetPosition2 x, y, 0#

            Set swTextFormat = swAnn.GetTextFormat(0)

            If Not swTextFormat Is Nothing Then
                swTextFormat.CharHeight = textSize
                swAnn.SetTextFormat 0, False, swTextFormat
            End If
        End If

    End If

    On Error GoTo 0

End Sub

Private Sub AddDrawingLine(ByVal swModel As Object, ByVal x1 As Double, ByVal y1 As Double, ByVal x2 As Double, ByVal y2 As Double)

    On Error Resume Next

    swModel.SketchManager.CreateLine x1, y1, 0#, x2, y2, 0#

    On Error GoTo 0

End Sub

Private Sub AddEndTicks(ByVal swModel As Object, ByVal x1 As Double, ByVal y1 As Double, ByVal x2 As Double, ByVal y2 As Double)

    Dim tick As Double

    tick = 0.005

    On Error Resume Next

    'Horizontal line ticks.
    If Abs(y2 - y1) < Abs(x2 - x1) Then
        swModel.SketchManager.CreateLine x1, y1 - tick, 0#, x1, y1 + tick, 0#
        swModel.SketchManager.CreateLine x2, y2 - tick, 0#, x2, y2 + tick, 0#
    Else
        'Vertical line ticks.
        swModel.SketchManager.CreateLine x1 - tick, y1, 0#, x1 + tick, y1, 0#
        swModel.SketchManager.CreateLine x2 - tick, y2, 0#, x2 + tick, y2, 0#
    End If

    On Error GoTo 0

End Sub
'===========================================================
' PER-PART DIMENSIONED DRAWINGS AND DWG EXPORTS
'
' Creates one SLDDRW, PDF and DWG for each generated part:
'   SB_001_MainBeam
'   SB_002_CrossMember
'   SB_003_DeckPanel
'   SB_004_GuardPost
'   SB_005_Handrail
'   SB_006_KickPlate
'   SB_007_EndPlate
'===========================================================

Public Sub CreatePartDrawings(ByRef d As BridgeDesign)

    Dim crossMemberY As Double
    Dim endPlateY As Double
    Dim endPlateZ As Double

    crossMemberY = d.BridgeWidthM + (2# * d.BeamWidthM) + 0.05
    endPlateY = d.BridgeWidthM + (2# * d.BeamWidthM) + SB_ENDPLATE_WIDTH_ALLOWANCE_M
    endPlateZ = d.BeamDepthM + SB_ENDPLATE_HEIGHT_ALLOWANCE_M

    CreateSinglePartDrawing d, "001", "Main beam", d.MainBeamPath, "2", d.SpanM, d.BeamWidthM, d.BeamDepthM
    CreateSinglePartDrawing d, "002", "Cross member", d.CrossMemberPath, CStr(d.CrossMemberCount), SB_CROSS_MEMBER_X_M, crossMemberY, SB_CROSS_MEMBER_Z_M
    CreateSinglePartDrawing d, "003", "Deck panel", d.DeckPanelPath, CStr(d.DeckPanelCount), d.DeckPanelActualLengthM, d.BridgeWidthM, SB_DECK_THICKNESS_M
    CreateSinglePartDrawing d, "004", "Guard post", d.GuardPostPath, CStr(d.GuardPostCount), SB_POST_SIZE_M, SB_POST_SIZE_M, d.RailHeightM
    CreateSinglePartDrawing d, "005", "Handrail", d.HandrailPath, "4", d.SpanM, SB_HANDRAIL_SIZE_M, SB_HANDRAIL_SIZE_M
    CreateSinglePartDrawing d, "006", "Kick plate", d.KickPlatePath, "2", d.SpanM, SB_KICKPLATE_THICKNESS_M, SB_KICKPLATE_HEIGHT_M
    CreateSinglePartDrawing d, "007", "End plate", d.EndPlatePath, "2", SB_ENDPLATE_THICKNESS_M, endPlateY, endPlateZ

End Sub

Private Sub CreateSinglePartDrawing(ByRef d As BridgeDesign, ByVal itemNo As String, ByVal partTitle As String, ByVal partPath As String, ByVal quantityText As String, ByVal sizeX As Double, ByVal sizeY As Double, ByVal sizeZ As Double)

    Dim swModel As Object
    Dim swDraw As Object
    Dim swSheet As Object
    Dim props As Variant
    Dim sheetW As Double
    Dim sheetH As Double
    Dim viewScale As Double
    Dim drawingPath As String
    Dim pdfPath As String
    Dim dwgPath As String
    Dim vFront As Object
    Dim vTop As Object
    Dim vRight As Object
    Dim vIso As Object

    OpenModelSilent partPath, SB_SW_DOC_PART

    Set swModel = NewModelDocument(SB_SW_DOC_DRAWING)
    Set swDraw = swModel

    sheetW = 0.42
    sheetH = 0.297

    On Error Resume Next
    Set swSheet = swDraw.GetCurrentSheet
    props = swSheet.GetProperties
    If IsArray(props) Then
        If UBound(props) >= 6 Then
            If CDbl(props(5)) > 0# Then sheetW = CDbl(props(5))
            If CDbl(props(6)) > 0# Then sheetH = CDbl(props(6))
        End If
    End If
    On Error GoTo 0

    viewScale = CalculateSafePartScale(sizeX, sizeY, sizeZ, sheetW, sheetH)

    Set vFront = swDraw.CreateDrawViewFromModelView3(partPath, "*Front", sheetW * 0.3, sheetH * 0.62, 0#)
    Set vTop = swDraw.CreateDrawViewFromModelView3(partPath, "*Top", sheetW * 0.3, sheetH * 0.32, 0#)
    Set vRight = swDraw.CreateDrawViewFromModelView3(partPath, "*Right", sheetW * 0.68, sheetH * 0.32, 0#)
    Set vIso = swDraw.CreateDrawViewFromModelView3(partPath, "*Isometric", sheetW * 0.68, sheetH * 0.66, 0#)

    ApplyViewScale vFront, viewScale
    ApplyViewScale vTop, viewScale
    ApplyViewScale vRight, viewScale
    ApplyViewScale vIso, viewScale * 0.85

    AddPartDimensionPack swModel, d, itemNo, partTitle, quantityText, sizeX, sizeY, sizeZ, sheetW, sheetH

    AddCustomPropertyText swModel, "Project", d.ProjectName
    AddCustomPropertyText swModel, "ItemNo", itemNo
    AddCustomPropertyText swModel, "PartTitle", partTitle
    AddCustomPropertyText swModel, "Quantity", quantityText
    AddCustomPropertyText swModel, "Material", d.MaterialName
    AddCustomPropertyText swModel, "Finish", d.FinishName
    AddCustomPropertyText swModel, "DrawingType", "Dimensioned part drawing"

    ZoomAndRebuild swModel

    drawingPath = PathCombine(d.OutputFolder, "SB_" & itemNo & "_" & CleanFileName(partTitle) & "_Drawing.SLDDRW")
    pdfPath = PathCombine(d.OutputFolder, "SB_" & itemNo & "_" & CleanFileName(partTitle) & "_Drawing.pdf")
    dwgPath = PathCombine(d.OutputFolder, "SB_" & itemNo & "_" & CleanFileName(partTitle) & "_Drawing.dwg")

    SaveModelAs swModel, drawingPath

    On Error Resume Next
    SaveModelAs swModel, pdfPath
    SaveModelAs swModel, dwgPath
    On Error GoTo 0

End Sub

Private Function CalculateSafePartScale(ByVal sizeX As Double, ByVal sizeY As Double, ByVal sizeZ As Double, ByVal sheetW As Double, ByVal sheetH As Double) As Double

    Dim availableW As Double
    Dim availableH As Double
    Dim maxLen As Double
    Dim maxHt As Double
    Dim s1 As Double
    Dim s2 As Double
    Dim s As Double

    availableW = sheetW * 0.34
    availableH = sheetH * 0.22

    maxLen = sizeX
    If sizeY > maxLen Then maxLen = sizeY

    maxHt = sizeZ
    If sizeY > maxHt Then maxHt = sizeY

    If maxLen <= 0# Then maxLen = 1#
    If maxHt <= 0# Then maxHt = 1#

    s1 = availableW / maxLen
    s2 = availableH / maxHt

    If s1 < s2 Then
        s = s1
    Else
        s = s2
    End If

    If s > 0.5 Then s = 0.5
    If s < 0.003 Then s = 0.003

    CalculateSafePartScale = s

End Function

Private Sub AddPartDimensionPack(ByVal swModel As Object, ByRef d As BridgeDesign, ByVal itemNo As String, ByVal partTitle As String, ByVal quantityText As String, ByVal sizeX As Double, ByVal sizeY As Double, ByVal sizeZ As Double, ByVal sheetW As Double, ByVal sheetH As Double)

    Dim textSize As Double
    Dim scheduleText As String

    textSize = 0.003

    AddDrawingNote swModel, "SMARTBRIDGE PART DRAWING - ITEM " & itemNo, sheetW * 0.04, sheetH * 0.975, textSize * 1.15
    AddDrawingNote swModel, partTitle, sheetW * 0.04, sheetH * 0.945, textSize

    AddDrawingNote swModel, "X LENGTH = " & FormatM(sizeX), sheetW * 0.16, sheetH * 0.5, textSize
    AddDrawingNote swModel, "Y WIDTH = " & FormatM(sizeY), sheetW * 0.16, sheetH * 0.2, textSize
    AddDrawingNote swModel, "Z HEIGHT = " & FormatM(sizeZ), sheetW * 0.62, sheetH * 0.2, textSize

    AddDrawingLine swModel, sheetW * 0.1, sheetH * 0.47, sheetW * 0.5, sheetH * 0.47
    AddDrawingLine swModel, sheetW * 0.1, sheetH * 0.17, sheetW * 0.5, sheetH * 0.17
    AddDrawingLine swModel, sheetW * 0.59, sheetH * 0.13, sheetW * 0.59, sheetH * 0.42

    AddEndTicks swModel, sheetW * 0.1, sheetH * 0.47, sheetW * 0.5, sheetH * 0.47
    AddEndTicks swModel, sheetW * 0.1, sheetH * 0.17, sheetW * 0.5, sheetH * 0.17
    AddEndTicks swModel, sheetW * 0.59, sheetH * 0.13, sheetW * 0.59, sheetH * 0.42

    scheduleText = "PART DIMENSION SCHEDULE" & vbCrLf
    scheduleText = scheduleText & "--------------------------" & vbCrLf
    scheduleText = scheduleText & "Item number: " & itemNo & vbCrLf
    scheduleText = scheduleText & "Part name: " & partTitle & vbCrLf
    scheduleText = scheduleText & "Quantity: " & quantityText & vbCrLf
    scheduleText = scheduleText & "X length: " & FormatM(sizeX) & vbCrLf
    scheduleText = scheduleText & "Y width: " & FormatM(sizeY) & vbCrLf
    scheduleText = scheduleText & "Z height: " & FormatM(sizeZ) & vbCrLf
    scheduleText = scheduleText & "Material: " & d.MaterialName & vbCrLf
    scheduleText = scheduleText & "Finish: " & d.FinishName & vbCrLf
    scheduleText = scheduleText & "Project: " & d.ProjectName

    AddDrawingNote swModel, scheduleText, sheetW * 0.56, sheetH * 0.88, textSize

End Sub
