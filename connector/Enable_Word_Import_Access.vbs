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

Dim Wsh, versions, v, root, roots, successCount, errCount

Set Wsh = CreateObject("WScript.Shell")
versions = Array("16.0", "15.0", "14.0", "12.0", "11.0")
roots = Array( _
    "HKCU\Software\Microsoft\Office\", _
    "HKCU\Software\Policies\Microsoft\Office\", _
    "HKLM\Software\Microsoft\Office\", _
    "HKLM\Software\Policies\Microsoft\Office\", _
    "HKLM\Software\WOW6432Node\Microsoft\Office\", _
    "HKLM\Software\WOW6432Node\Policies\Microsoft\Office\" _
)
successCount = 0
errCount = 0

For Each v In versions
    For Each root In roots
        On Error Resume Next
        Wsh.RegWrite root & v & "\Word\Security\AccessVBOM", 1, "REG_DWORD"
        Wsh.RegWrite root & v & "\Word\Security\Level", 1, "REG_DWORD"
        Wsh.RegWrite root & v & "\Word\Security\VBAWarnings", 1, "REG_DWORD"
        Wsh.RegWrite root & v & "\Word\Security\DisableAllMacros", 0, "REG_DWORD"
        Wsh.RegWrite root & v & "\Word\Security\ExtensionHardening", 0, "REG_DWORD"
        Wsh.RegWrite root & v & "\Word\Options\AccessVBOM", 1, "REG_DWORD"
        Wsh.RegWrite root & v & "\Common\Security\AccessVBOM", 1, "REG_DWORD"
        If Err.Number = 0 Then
            successCount = successCount + 1
        Else
            errCount = errCount + 1
        End If
        On Error GoTo 0
    Next
Next

MsgBox "Microsoft Word Import Access Granted!" & vbCrLf & vbCrLf & _
       "• Trust access to VBA project object model: ENABLED" & vbCrLf & _
       "• Macro execution security level: ENABLE ALL" & vbCrLf & _
       "• VBA macro warnings: DISABLED (Full Trust)" & vbCrLf & vbCrLf & _
       "Configured for " & successCount & " Microsoft Office version registry keys." & vbCrLf & _
       "If Word is currently open, please restart Word to apply the new permissions.", _
       vbInformation, "Word Toolkit - Access Manager"
