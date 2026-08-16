Attribute VB_Name = "Toolkit_Shortcuts"
Option Explicit

Sub ExportKeyboardShortcuts()
    Dim kb As Object, filePath As Variant, fnum As Integer, count As Integer
    On Error GoTo HandleErr

    filePath = Application.GetSaveAsFilename(InitialFileName:="WordKeyboardShortcuts.csv", FileFilter:="CSV Files (*.csv), *.csv")
    If filePath = False Or filePath = "False" Then Exit Sub

    fnum = FreeFile
    Open filePath For Output As #fnum
    Print #fnum, "KeyCategory,Command,KeyCode,KeyCode2,KeyString"

    For Each kb In Application.KeyBindings
        Print #fnum, kb.KeyCategory & "," & kb.Command & "," & kb.KeyCode & "," & kb.KeyCode2 & "," & Chr(34) & kb.KeyString & Chr(34)
        count = count + 1
    Next kb

    Close #fnum
    MsgBox count & " keyboard shortcut(s) exported to:" & vbCrLf & filePath, vbInformation, "Export Shortcuts Successful"
    Exit Sub

HandleErr:
    If fnum <> 0 Then Close #fnum
    MsgBox "Could not export keyboard shortcuts." & vbCrLf & "Error: " & Err.Description, vbCritical, "Export Shortcuts Error"
End Sub

Sub ImportKeyboardShortcuts()
    Dim filePath As Variant, fnum As Integer, ln As String, parts() As String, count As Integer, failCount As Integer
    On Error GoTo HandleErr

    filePath = Application.GetOpenFilename("CSV Files (*.csv), *.csv")
    If filePath = False Or filePath = "False" Then Exit Sub

    fnum = FreeFile
    Open filePath For Input As #fnum
    If Not EOF(fnum) Then Line Input #fnum, ln ' skip header

    Application.CustomizationContext = NormalTemplate

    Do While Not EOF(fnum)
        Line Input #fnum, ln
        If Len(Trim(ln)) > 0 Then
            parts = Split(ln, ",")
            If UBound(parts) >= 3 Then
                On Error Resume Next
                Err.Clear
                Application.KeyBindings.Add KeyCategory:=CLng(parts(0)), Command:=parts(1), KeyCode:=CLng(parts(2)), KeyCode2:=IIf(parts(3) = "0" Or parts(3) = "", 0, CLng(parts(3)))
                If Err.Number = 0 Then count = count + 1 Else failCount = failCount + 1
                On Error GoTo HandleErr
            End If
        End If
    Loop

    Close #fnum
    MsgBox count & " shortcut(s) imported successfully." & IIf(failCount > 0, vbCrLf & failCount & " could not be applied.", ""), vbInformation, "Import Shortcuts Successful"
    Exit Sub

HandleErr:
    If fnum <> 0 Then Close #fnum
    MsgBox "Could not import keyboard shortcuts." & vbCrLf & "Error: " & Err.Description, vbCritical, "Import Shortcuts Error"
End Sub

Sub ResetAllKeyboardShortcuts()
    If MsgBox("This will reset ALL custom keyboard shortcuts to Word defaults. Continue?", vbYesNo + vbExclamation, "Reset Shortcuts") = vbYes Then
        Application.CustomizationContext = NormalTemplate
        Application.KeyBindings.ClearAll
        MsgBox "Keyboard shortcuts reset to Word defaults.", vbInformation, "Reset Shortcuts"
    End If
End Sub
