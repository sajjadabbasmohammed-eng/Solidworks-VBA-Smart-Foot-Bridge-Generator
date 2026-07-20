Attribute VB_Name = "modSmartBridgeSolidWorksCore"
Option Explicit

Public Function ConnectToSolidWorks() As Boolean

    On Error Resume Next
    Set g_swApp = Application.SldWorks
    If g_swApp Is Nothing Then Set g_swApp = GetObject(, "SldWorks.Application")
    On Error GoTo 0

    If g_swApp Is Nothing Then
        MsgBox "Could not connect to SolidWorks. Open SolidWorks and run the macro again.", vbCritical, "SmartBridge API"
        ConnectToSolidWorks = False
    Else
        g_swApp.Visible = True
        ConnectToSolidWorks = True
    End If

End Function

Public Function NewModelDocument(ByVal docType As Long) As Object

    Dim templatePath As String

    templatePath = GetDefaultTemplatePath(docType)
    If Len(templatePath) = 0 Then Err.Raise vbObjectError + 1010, "NewModelDocument", "No default template path was found."
    If Not FileExistsSB(templatePath) Then Err.Raise vbObjectError + 1011, "NewModelDocument", "Template file was not found:" & vbCrLf & templatePath

    Set NewModelDocument = g_swApp.NewDocument(templatePath, 0, 0#, 0#)
    If NewModelDocument Is Nothing Then Err.Raise vbObjectError + 1012, "NewModelDocument", "SolidWorks failed to create a new document from:" & vbCrLf & templatePath

End Function

Public Function GetDefaultTemplatePath(ByVal docType As Long) As String

    Dim pref As Long
    Dim p As String

    Select Case docType
        Case SB_SW_DOC_PART
            pref = SB_SW_DEFAULT_TEMPLATE_PART
        Case SB_SW_DOC_ASSEMBLY
            pref = SB_SW_DEFAULT_TEMPLATE_ASSEMBLY
        Case SB_SW_DOC_DRAWING
            pref = SB_SW_DEFAULT_TEMPLATE_DRAWING
        Case Else
            Err.Raise vbObjectError + 1013, "GetDefaultTemplatePath", "Unknown document type."
    End Select

    On Error Resume Next
    p = CStr(g_swApp.GetUserPreferenceStringValue(pref))
    On Error GoTo 0

    If Len(p) > 0 Then
        If FileExistsSB(p) Then
            GetDefaultTemplatePath = NormalizeFilePath(p)
            Exit Function
        End If
    End If

    p = FindCommonSolidWorksTemplate(docType)
    If Len(p) > 0 Then
        GetDefaultTemplatePath = p
        Exit Function
    End If

    Err.Raise vbObjectError + 1014, "GetDefaultTemplatePath", "No default template was found. Set default templates in SolidWorks."

End Function

Private Function FindCommonSolidWorksTemplate(ByVal docType As Long) As String

    Dim fileName As String
    Dim p As String
    Dim y As Long

    Select Case docType
        Case SB_SW_DOC_PART
            fileName = "Part.prtdot"
        Case SB_SW_DOC_ASSEMBLY
            fileName = "Assembly.asmdot"
        Case SB_SW_DOC_DRAWING
            fileName = "Drawing.drwdot"
        Case Else
            FindCommonSolidWorksTemplate = ""
            Exit Function
    End Select

    For y = 2026 To 2020 Step -1
        p = "C:\ProgramData\SOLIDWORKS\SOLIDWORKS " & y & "\templates\" & fileName
        If FileExistsSB(p) Then
            FindCommonSolidWorksTemplate = p
            Exit Function
        End If
    Next y

    FindCommonSolidWorksTemplate = ""

End Function

Public Sub SaveModelAs(ByVal swModel As Object, ByVal filePath As String)

    Dim errors As Long
    Dim warnings As Long
    Dim ok As Boolean
    Dim p As String
    Dim fso As Object
    Dim parentFolder As String

    If swModel Is Nothing Then Err.Raise vbObjectError + 1020, "SaveModelAs", "Nothing was supplied as the SolidWorks model."

    p = NormalizeFilePath(filePath)
    Set fso = CreateObject("Scripting.FileSystemObject")
    parentFolder = fso.GetParentFolderName(p)
    If Len(parentFolder) > 0 Then EnsureFolderExists parentFolder

    swModel.ClearSelection2 True
    ok = swModel.Extension.SaveAs(p, SB_SW_SAVE_AS_CURRENT_VERSION, SB_SW_SAVE_SILENT, Nothing, errors, warnings)

    If Not ok Then
        Err.Raise vbObjectError + 1021, "SaveModelAs", "SolidWorks could not save:" & vbCrLf & p & vbCrLf & "Error code: " & errors & vbCrLf & "Warning code: " & warnings
    End If

End Sub

Public Function OpenModelSilent(ByVal filePath As String, ByVal docType As Long) As Object

    Dim errors As Long
    Dim warnings As Long
    Dim p As String
    Dim alreadyOpen As Object

    p = NormalizeFilePath(filePath)

    If Not FileExistsSB(p) Then Err.Raise vbObjectError + 1030, "OpenModelSilent", "File not found:" & vbCrLf & p

    On Error Resume Next
    Set alreadyOpen = g_swApp.GetOpenDocumentByName(p)
    On Error GoTo 0

    If Not alreadyOpen Is Nothing Then
        Set OpenModelSilent = alreadyOpen
        Exit Function
    End If

    Set OpenModelSilent = g_swApp.OpenDoc6(p, docType, SB_SW_OPEN_SILENT, "", errors, warnings)

    If OpenModelSilent Is Nothing Then
        Err.Raise vbObjectError + 1031, "OpenModelSilent", "SolidWorks could not open:" & vbCrLf & p & vbCrLf & "Error code: " & errors & vbCrLf & "Warning code: " & warnings
    End If

End Function

Public Sub AddCustomPropertyText(ByVal swModel As Object, ByVal propName As String, ByVal propValue As String)

    Dim propMgr As Object

    On Error Resume Next
    Set propMgr = swModel.Extension.CustomPropertyManager("")
    If Not propMgr Is Nothing Then propMgr.Add3 propName, SB_SW_CUSTOMINFO_TEXT, propValue, SB_SW_CUSTOMPROP_REPLACE
    On Error GoTo 0

End Sub

Public Sub ApplyAppearanceColor(ByVal swModel As Object, ByVal rVal As Long, ByVal gVal As Long, ByVal bVal As Long)

    Dim matProps(8) As Double

    matProps(0) = rVal / 255#
    matProps(1) = gVal / 255#
    matProps(2) = bVal / 255#
    matProps(3) = 0.9
    matProps(4) = 0.75
    matProps(5) = 0.35
    matProps(6) = 0.25
    matProps(7) = 0#
    matProps(8) = 0#

    On Error Resume Next
    swModel.MaterialPropertyValues = matProps
    swModel.GraphicsRedraw2
    On Error GoTo 0

End Sub

Public Sub TryApplySolidWorksMaterial(ByVal swModel As Object, ByVal MaterialName As String)

    On Error Resume Next
    swModel.SetMaterialPropertyName2 "", "", MaterialName
    On Error GoTo 0

End Sub

Public Sub ZoomAndRebuild(ByVal swModel As Object)

    On Error Resume Next
    swModel.ForceRebuild3 False
    swModel.ViewZoomtofit2
    swModel.GraphicsRedraw2
    On Error GoTo 0

End Sub
