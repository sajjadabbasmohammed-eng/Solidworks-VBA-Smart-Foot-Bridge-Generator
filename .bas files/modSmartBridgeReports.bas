Attribute VB_Name = "modSmartBridgeReports"
Option Explicit

Public Sub WriteBridgeBomCsv(ByRef d As BridgeDesign)

    Dim f As Integer

    f = FreeFile
    Open d.BomCsvPath For Output As #f

    Print #f, CsvLine("Item", "Part", "File", "Quantity", "Key dimension", "Material/Finish")
    Print #f, CsvLine("001", "Main beam", d.MainBeamPath, "2", FormatM(d.SpanM), d.MaterialName & " / " & d.FinishName)
    Print #f, CsvLine("002", "Cross member", d.CrossMemberPath, CStr(d.CrossMemberCount), FormatM(d.BridgeWidthM), d.MaterialName & " / " & d.FinishName)
    Print #f, CsvLine("003", "Deck panel", d.DeckPanelPath, CStr(d.DeckPanelCount), FormatM(d.DeckPanelActualLengthM), d.DeckName & " / " & d.FinishName)
    Print #f, CsvLine("004", "Guard post", d.GuardPostPath, CStr(d.GuardPostCount), FormatM(d.RailHeightM), d.MaterialName & " / " & d.FinishName)
    Print #f, CsvLine("005", "Handrail", d.HandrailPath, "4", FormatM(d.SpanM), d.MaterialName & " / " & d.FinishName)
    Print #f, CsvLine("006", "Kick plate", d.KickPlatePath, "2", FormatM(d.SpanM), d.MaterialName & " / " & d.FinishName)
    Print #f, CsvLine("007", "End plate", d.EndPlatePath, "2", FormatMM(SB_ENDPLATE_THICKNESS_M), d.MaterialName & " / " & d.FinishName)
    Print #f, CsvLine("", "", "", "", "", "")
    Print #f, CsvLine("SUMMARY", "Estimated mass", "", "", Format$(d.EstimatedMassKG, "0") & " kg", "")
    Print #f, CsvLine("SUMMARY", "Estimated cost", "", "", "GBP " & Format$(d.EstimatedCostGBP, "0"), "")

    Close #f

End Sub

Public Sub WriteBridgeReportTxt(ByRef d As BridgeDesign)

    Dim f As Integer

    f = FreeFile
    Open d.ReportPath For Output As #f

    Print #f, "SMARTBRIDGE API - AUTOMATIC DESIGN REPORT"
    Print #f, String$(70, "=")
    Print #f, "Generated on: " & d.CreatedOn
    Print #f, "Project name: " & d.ProjectName
    Print #f, "Output folder: " & d.OutputFolder
    Print #f, "Data workbook: " & d.DataWorkbookPath
    Print #f, ""
    Print #f, "VALIDATED INPUTS"
    Print #f, String$(70, "-")
    Print #f, "Span: " & FormatM(d.SpanM)
    Print #f, "Bridge width: " & FormatM(d.BridgeWidthM)
    Print #f, "Material: " & d.MaterialName
    Print #f, "Deck: " & d.DeckName
    Print #f, "Finish: " & d.FinishName
    Print #f, "Finish RGB: " & d.finishR & "," & d.finishG & "," & d.finishB
    Print #f, "Finish multiplier: " & Format$(d.FinishMultiplier, "0.000")
    Print #f, ""
    Print #f, "SELECTED STRUCTURAL MEMBER"
    Print #f, String$(70, "-")
    Print #f, "Selected beam: " & d.BeamName
    Print #f, "Beam depth: " & FormatMM(d.BeamDepthM)
    Print #f, "Beam width: " & FormatMM(d.BeamWidthM)
    Print #f, "Beam area: " & Format$(d.BeamAreaM2 * 1000000#, "0.0") & " mm2"
    Print #f, ""
    Print #f, "AUTOMATIC LAYOUT"
    Print #f, String$(70, "-")
    Print #f, "Cross member count: " & d.CrossMemberCount
    Print #f, "Actual cross member spacing: " & FormatMM(d.ActualCrossSpacingM)
    Print #f, "Deck panel count: " & d.DeckPanelCount
    Print #f, "Actual deck panel length: " & FormatM(d.DeckPanelActualLengthM)
    Print #f, "Guard post count: " & d.GuardPostCount
    Print #f, "Rail height: " & FormatM(d.RailHeightM)
    Print #f, ""
    Print #f, "STRUCTURAL EQUATIONS USED"
    Print #f, String$(70, "-")
    Print #f, "Mmax = wL^2 / 8"
    Print #f, "Vmax = wL / 2"
    Print #f, "sigma = M / Z"
    Print #f, "tau = 1.5V / A"
    Print #f, "sigmaVM = sqrt(sigma^2 + 3tau^2)"
    Print #f, "delta = 5wL^4 / 384EI"
    Print #f, "f1 = pi/2 * sqrt(EI / mL^4)"
    Print #f, ""
    Print #f, "ESTIMATE SUMMARY"
    Print #f, String$(70, "-")
    Print #f, "Estimated mass: " & Format$(d.EstimatedMassKG, "0") & " kg"
    Print #f, "Estimated cost: GBP " & Format$(d.EstimatedCostGBP, "0")
    Print #f, ""
    Print #f, "GENERATED FILES"
    Print #f, String$(70, "-")
    Print #f, "Main beam part: " & d.MainBeamPath
    Print #f, "Cross member part: " & d.CrossMemberPath
    Print #f, "Deck panel part: " & d.DeckPanelPath
    Print #f, "Guard post part: " & d.GuardPostPath
    Print #f, "Handrail part: " & d.HandrailPath
    Print #f, "Kick plate part: " & d.KickPlatePath
    Print #f, "End plate part: " & d.EndPlatePath
    Print #f, "Assembly: " & d.AssemblyPath
    Print #f, "Drawing: " & d.drawingPath
    Print #f, "PDF drawing: " & d.DrawingPdfPath
    Print #f, "DWG drawing: " & d.DrawingDwgPath
    Print #f, "CSV BOM: " & d.BomCsvPath

    Close #f

End Sub

Private Function CsvLine(ByVal a As String, ByVal b As String, ByVal c As String, ByVal d As String, ByVal e As String, ByVal f As String) As String
    CsvLine = CsvQ(a) & "," & CsvQ(b) & "," & CsvQ(c) & "," & CsvQ(d) & "," & CsvQ(e) & "," & CsvQ(f)
End Function

Private Function CsvQ(ByVal rawText As String) As String
    CsvQ = """" & Replace(rawText, """", """""") & """"
End Function
