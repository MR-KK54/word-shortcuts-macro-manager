Attribute VB_Name = "Toolkit_Macros"
Option Explicit

Const vbext_ct_StdModule = 1
Const vbext_ct_ClassModule = 2
Const vbext_ct_MSForm = 3

' Export all macro modules from active Word VBA project to a folder
Sub ExportAllMacros()
    Dim vbProj As Object, vbComp As Object
    Dim exportPath As String, ext As String, count As Integer
    On Error GoTo HandleErr

    exportPath = GetFolder("Select a folder to export macro modules to")
    If exportPath = "" Then Exit Sub

    Set vbProj = Application.VBE.ActiveVBProject
    For Each vbComp In vbProj.VBComponents
        Select Case vbComp.Type
            Case vbext_ct_StdModule: ext = ".bas"
            Case vbext_ct_ClassModule: ext = ".cls"
            Case vbext_ct_MSForm: ext = ".frm"
            Case Else: ext = ""
        End Select

        ' Do not export built-in document modules (e.g. ThisDocument) as standalone .bas
        If ext <> "" And vbComp.Name <> "ThisDocument" Then
            vbComp.Export exportPath & "\" & vbComp.Name & ext
            count = count + 1
        End If
    Next vbComp

    MsgBox count & " macro module(s) exported to:" & vbCrLf & exportPath, vbInformation, "Export Macros Successful"
    Exit Sub

HandleErr:
    MsgBox "Could not export macros. Make sure 'Trust access to the VBA project object model' is enabled." & vbCrLf & "Error: " & Err.Description, vbCritical, "Export Macros Error"
End Sub

' Import macro modules from a folder into Word.
' REQUIREMENT: If the macro module already exists in Word, replace it!
Sub ImportAllMacros()
    Dim vbProj As Object, importPath As String
    Dim fso As Object, fld As Object, f As Object
    Dim compName As String, ext As String, count As Integer, replacedCount As Integer
    Dim existingComp As Object

    On Error GoTo HandleErr
    importPath = GetFolder("Select the folder containing .bas / .cls / .frm files to import")
    If importPath = "" Then Exit Sub

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set vbProj = Application.VBE.ActiveVBProject
    Set fld = fso.GetFolder(importPath)

    For Each f In fld.Files
        ext = LCase(fso.GetExtensionName(f.Name))
        If ext = "bas" Or ext = "cls" Or ext = "frm" Then
            compName = fso.GetBaseName(f.Name)

            ' --- REQUIREMENT: Replace existing macro if already in Word ---
            On Error Resume Next
            Set existingComp = Nothing
            Set existingComp = vbProj.VBComponents(compName)
            If Not existingComp Is Nothing Then
                vbProj.VBComponents.Remove existingComp
                replacedCount = replacedCount + 1
            End If
            On Error GoTo HandleErr

            ' Import new module version into Word
            vbProj.VBComponents.Import f.Path
            count = count + 1
        End If
    Next f

    MsgBox count & " macro module(s) imported (" & replacedCount & " existing module(s) replaced)." & vbCrLf & _
           "Remember to save Normal.dotm or your active template.", vbInformation, "Import Macros Successful"
    Exit Sub

HandleErr:
    MsgBox "Could not import macros. Make sure 'Trust access to the VBA project object model' is enabled." & vbCrLf & "Error: " & Err.Description, vbCritical, "Import Macros Error"
End Sub
