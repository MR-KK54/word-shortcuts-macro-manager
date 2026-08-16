Attribute VB_Name = "Toolkit_RibbonQAT"
Option Explicit

Private Function OfficeUIPath() As String
    OfficeUIPath = Environ$("LOCALAPPDATA") & "\Microsoft\Office\Word.officeUI"
End Function

Sub ExportRibbonCustomizations()
    Dim srcFile As String, destFolder As String, fso As Object
    srcFile = OfficeUIPath()
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(srcFile) Then
        MsgBox "No custom ribbon/QAT file was found at:" & vbCrLf & srcFile & vbCrLf & vbCrLf & "Use File > Options > Customize Ribbon > Import/Export in Word instead.", vbExclamation, "Export Ribbon/QAT"
        Exit Sub
    End If
    destFolder = GetFolder("Select a folder to export the ribbon/QAT customization to")
    If destFolder = "" Then Exit Sub
    On Error GoTo HandleErr
    fso.CopyFile srcFile, destFolder & "\Word.officeUI", True
    MsgBox "Ribbon and Quick Access Toolbar customizations exported to:" & vbCrLf & destFolder, vbInformation, "Export Ribbon/QAT"
    Exit Sub
HandleErr:
    MsgBox "Could not export the ribbon/QAT file." & vbCrLf & "Error: " & Err.Description, vbCritical, "Export Ribbon/QAT Error"
End Sub

Sub ImportRibbonCustomizations()
    Dim srcFile As Variant, destFile As String, fso As Object
    srcFile = Application.GetOpenFilename("Office UI Files (*.officeUI), *.officeUI")
    If srcFile = False Or srcFile = "False" Then Exit Sub
    destFile = OfficeUIPath()
    Set fso = CreateObject("Scripting.FileSystemObject")
    On Error GoTo HandleErr
    fso.CopyFile srcFile, destFile, True
    MsgBox "Ribbon/QAT customization imported. Close and reopen Word for changes to take effect.", vbInformation, "Import Ribbon/QAT"
    Exit Sub
HandleErr:
    MsgBox "Could not import the ribbon/QAT file. If Word reports it is in use, close Word and copy it manually to:" & vbCrLf & destFile & vbCrLf & vbCrLf & "Error: " & Err.Description, vbCritical, "Import Ribbon/QAT Error"
End Sub
