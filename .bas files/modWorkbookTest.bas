Attribute VB_Name = "modWorkbookTest"
Option Explicit

Public Sub TestSmartBridgeWorkbook()

    Dim p As String
    Dim xlApp As Object
    Dim wb As Object
    Dim ws As Object
    Dim msg As String

    p = "J:\api 206\SmartBridgeSections\SmartBridgeData.xlsx"
    p = NormalizeFilePath(p)

    If Not FileExistsSB(p) Then
        MsgBox "Workbook not found:" & vbCrLf & p, vbExclamation, "Workbook test"
        Exit Sub
    End If

    On Error GoTo FailTest

    Set xlApp = CreateObject("Excel.Application")
    xlApp.Visible = False
    Set wb = xlApp.Workbooks.Open(p, False, True)

    msg = "Workbook opened successfully:" & vbCrLf
    msg = msg & p & vbCrLf & vbCrLf
    msg = msg & "Sheets found:" & vbCrLf

    For Each ws In wb.Worksheets
        msg = msg & "- [" & ws.Name & "]" & vbCrLf
    Next ws

    wb.Close False
    xlApp.Quit

    Set wb = Nothing
    Set xlApp = Nothing

    MsgBox msg, vbInformation, "Workbook test"
    Exit Sub

FailTest:

    On Error Resume Next
    If Not wb Is Nothing Then wb.Close False
    If Not xlApp Is Nothing Then xlApp.Quit

    MsgBox "Workbook test failed." & vbCrLf & vbCrLf & _
           "Path:" & vbCrLf & p & vbCrLf & vbCrLf & _
           "Error " & Err.Number & ": " & Err.description, _
           vbExclamation, "Workbook test"

End Sub

