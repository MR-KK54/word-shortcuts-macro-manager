Option Explicit
' ============================================================
'  WORD TOOLKIT - ENABLE ALL WORD IMPORT ACCESS (standalone VBS)
'  Run this file anytime to grant Microsoft Word full import
'  permissions without security blocks or warnings.
'
'  What it configures in Windows Registry for MS Word:
'   - AccessVBOM = 1  ("Trust access to VBA project object model")
'   - Level = 1       ("Enable all macros")
'   - VBAWarnings = 1 (No warnings for unsigned macros)
'   - DisableAllMacros = 0
'   - ExtensionHardening = 0
' ============================================================

Dim Wsh, versions, v, secKey, successCount, errCount

Set Wsh = CreateObject("WScript.Shell")
versions = Array("16.0", "15.0", "14.0", "12.0", "11.0")
successCount = 0
errCount = 0

For Each v In versions
    secKey = "HKCU\Software\Microsoft\Office\" & v & "\Word\Security\"
    On Error Resume Next
    Wsh.RegWrite secKey & "AccessVBOM", 1, "REG_DWORD"
    Wsh.RegWrite secKey & "Level", 1, "REG_DWORD"
    Wsh.RegWrite secKey & "VBAWarnings", 1, "REG_DWORD"
    Wsh.RegWrite secKey & "DisableAllMacros", 0, "REG_DWORD"
    Wsh.RegWrite secKey & "ExtensionHardening", 0, "REG_DWORD"
    If Err.Number = 0 Then
        successCount = successCount + 1
    Else
        errCount = errCount + 1
    End If
    On Error GoTo 0
Next

MsgBox "Microsoft Word Import Access Granted!" & vbCrLf & vbCrLf & _
       "• Trust access to VBA project object model: ENABLED" & vbCrLf & _
       "• Macro execution security level: ENABLE ALL" & vbCrLf & _
       "• VBA macro warnings: DISABLED (Full Trust)" & vbCrLf & vbCrLf & _
       "Configured for " & successCount & " Microsoft Office version registry keys." & vbCrLf & _
       "If Word is currently open, please restart Word to apply the new permissions.", _
       vbInformation, "Word Toolkit - Access Manager"
