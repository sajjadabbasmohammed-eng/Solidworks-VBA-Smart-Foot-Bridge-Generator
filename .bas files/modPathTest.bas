Attribute VB_Name = "modPathTest"
Option Explicit

Public Sub TestSmartBridgePaths()

    Dim csvPath As String
    Dim outPath As String
    Dim msg As String

    csvPath = "J:\api 206\SmartBridgeSections\SmartBridgeSections.csv"
    outPath = "J:\api 206\Output"

    msg = "CSV path:" & vbCrLf
    msg = msg & NormalizeFilePath(csvPath) & vbCrLf & vbCrLf
    msg = msg & "CSV exists: " & CStr(FileExistsSB(csvPath)) & vbCrLf & vbCrLf

    msg = msg & "Output folder:" & vbCrLf
    msg = msg & NormalizeFolderPath(outPath) & vbCrLf & vbCrLf

    EnsureFolderExists outPath

    msg = msg & "Output folder check completed."

    MsgBox msg, vbInformation, "SmartBridge path test"

End Sub
