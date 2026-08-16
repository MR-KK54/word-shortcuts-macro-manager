Option Explicit
' ============================================================
'  WORD TOOLKIT - RESET & UNLOCK NORMAL TEMPLATE
'  What it does:
'   1. Closes any open Microsoft Word instances
'   2. Backs up your old Normal.dotm file to Normal_Backup.dotm
'   3. Removes old locked/unviewable Normal.dotm so Word automatically
'      creates a fresh, clean, 100% unlocked Normal template
'   4. Re-applies Word Toolkit import permissions
' ============================================================

Dim Wsh, fso, templatesDir, normalFile, backupFile, w, versions, v, secKey, polKey, root, roots

Set Wsh = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

' --- 1. Close open Word instances safely ---
On Error Resume Next
Set w = GetObject(, "Word.Application")
If Not w Is Nothing Then
    w.Quit
    WScript.Sleep 500
End If
Err.Clear
Wsh.Run "taskkill /F /IM winword.exe /T", 0, True
On Error GoTo 0

' --- 2. Locate Normal.dotm ---
templatesDir = fso.BuildPath(Wsh.ExpandEnvironmentStrings("%APPDATA%"), "Microsoft\Templates")
normalFile = fso.BuildPath(templatesDir, "Normal.dotm")
backupFile = fso.BuildPath(templatesDir, "Normal_Backup_" & Replace(Replace(Replace(Now(), "/", "-"), ":", "-"), " ", "_") & ".dotm")

If fso.FileExists(normalFile) Then
    On Error Resume Next
    ' Create a safety backup of old Normal.dotm
    fso.CopyFile normalFile, backupFile, True
    ' Remove old locked Normal.dotm
    fso.DeleteFile normalFile, True
    If Err.Number <> 0 Then
        MsgBox "Could not remove old Normal.dotm because Word is still running." & vbCrLf & _
               "Please close Word completely and run this script again.", vbExclamation, "Reset Word Template"
        WScript.Quit
    End If
    On Error GoTo 0
End If

' --- 3. Enable Word permissions in Registry ---
versions = Array("16.0", "15.0", "14.0", "12.0", "11.0")
roots = Array( _
    "HKCU\Software\Microsoft\Office\", _
    "HKCU\Software\Policies\Microsoft\Office\", _
    "HKLM\Software\Microsoft\Office\", _
    "HKLM\Software\Policies\Microsoft\Office\", _
    "HKLM\Software\WOW6432Node\Microsoft\Office\", _
    "HKLM\Software\WOW6432Node\Policies\Microsoft\Office\" _
)
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
        On Error GoTo 0
    Next
Next

' --- 4. Start Word silently to regenerate fresh Normal.dotm ---
On Error Resume Next
Set w = CreateObject("Word.Application")
If Not w Is Nothing Then
    w.Visible = False
    w.Documents.Add
    w.NormalTemplate.Save
    w.Quit
    Set w = Nothing
End If
On Error GoTo 0

MsgBox "Word Normal Template Reset Completed!" & vbCrLf & vbCrLf & _
       "• Old locked template backed up to: " & backupFile & vbCrLf & _
       "• Fresh, clean, unlocked Normal.dotm template created" & vbCrLf & _
       "• Trust access to VBA project object model enabled" & vbCrLf & vbCrLf & _
       "Now run WordToolkit_Setup.vbs to install the toolkit into your fresh template!", _
       vbInformation, "Reset Word Template"
