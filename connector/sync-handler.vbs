Option Explicit
' ============================================================
' WORD TOOLKIT - sync handler (receives wordtoolkit://sync links)
'  - Enables Word's import access automatically every time
'  - Self-updates from the server when a "u" (server URL) is given
'  - Tells the running Microsoft Word to install the selection
' This file is downloaded by WordToolkit_Setup.vbs and can also
' update itself on every click, so it never goes stale.
' ============================================================

Dim argLine, parts, kv, i, dict, p, k, u, m, s, r, w, Wsh, v, secKey
Dim here, fso, f, http, own, scriptDir, syncComp, compsContainer, appDir, basFiles, bFile, bName, cObj

Set Wsh = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")
scriptDir = fso.GetParentFolderName(WScript.ScriptFullName)
here = WScript.ScriptFullName

argLine = ""
If WScript.Arguments.Count > 0 Then argLine = WScript.Arguments(0)

Set dict = CreateObject("Scripting.Dictionary")
parts = Split(argLine, "&")
For i = 0 To UBound(parts)
    kv = Split(parts(i), "=")
    If UBound(kv) >= 1 Then
        p = kv(0)
        k = Mid(p, InStrRev(p, "?") + 1)
        dict(k) = Unescape(kv(1))
    End If
Next

' --- Enable Word import access automatically (Trust access + all macros) ---
For Each v In Array("16.0", "15.0", "14.0", "12.0", "11.0")
    secKey = "HKCU\Software\Microsoft\Office\" & v & "\Word\Security\"
    On Error Resume Next
    Wsh.RegWrite secKey & "AccessVBOM", 1, "REG_DWORD"
    Wsh.RegWrite secKey & "Level", 1, "REG_DWORD"
    Wsh.RegWrite secKey & "VBAWarnings", 1, "REG_DWORD"
    Wsh.RegWrite secKey & "DisableAllMacros", 0, "REG_DWORD"
    Wsh.RegWrite secKey & "ExtensionHardening", 0, "REG_DWORD"
    On Error GoTo 0
Next

' --- Self-update: fetch the newest handler from the server and replace this file ---
If dict.Exists("u") And dict("u") <> "" Then
    Set http = CreateObject("MSXML2.XMLHTTP")
    On Error Resume Next
    http.Open "GET", dict("u") & "/connector/sync-handler.vbs", False
    http.Send
    If Err.Number = 0 And http.Status >= 200 And http.Status < 300 Then
        own = ""
        On Error Resume Next
        Set f = fso.OpenTextFile(here, 1, False, True)
        If Not f Is Nothing Then
            own = f.ReadAll
            f.Close
        End If
        On Error GoTo 0
        If StrComp(own, http.ResponseText, 1) <> 0 Then
            Set f = fso.CreateTextFile(here, True, True)
            f.Write http.ResponseText
            f.Close
            ' Re-run the freshly updated handler with the same arguments
            Wsh.Run """" & WScript.FullName & """" & " """ & here & """ """ & argLine & """", 0, False
            WScript.Quit
        End If
    End If
    On Error GoTo 0
End If

' --- Special action: only enable access (web tool button) ---
If dict.Exists("act") And dict("act") = "enable" Then
    MsgBox "Word import access enabled (Trust access + all macros)." & vbCrLf & _
           "Restart Word once to apply the new permissions, then install again." & vbCrLf & _
           "Nothing was imported.", vbInformation, "Word Toolkit"
    WScript.Quit
End If

' --- Direct install: tell Word to fetch the selection from the server ---
u = dict("u"): m = dict("m"): s = dict("s"): r = dict("r")

On Error Resume Next
Set w = GetObject(, "Word.Application")
If w Is Nothing Then
    Set w = CreateObject("Word.Application")
    w.Visible = True
End If
If Err.Number <> 0 Or w Is Nothing Then
    MsgBox "Could not start Word: " & Err.Description, vbCritical, "Word Toolkit"
    WScript.Quit
End If
On Error GoTo 0

If w.Documents.Count = 0 Then
    On Error Resume Next
    w.Documents.Add
    On Error GoTo 0
End If

' Ensure Toolkit_Sync is imported into Word before executing
Set syncComp = Nothing
On Error Resume Next
Set syncComp = w.NormalTemplate.VBProject.VBComponents("Toolkit_Sync")
If syncComp Is Nothing And w.Documents.Count > 0 Then
    Set syncComp = w.ActiveDocument.VBProject.VBComponents("Toolkit_Sync")
End If
On Error GoTo 0

If syncComp Is Nothing Then
    appDir = fso.BuildPath(Wsh.ExpandEnvironmentStrings("%APPDATA%"), "WordToolkit")
    basFiles = Array("Toolkit_Helpers.bas", "Toolkit_Macros.bas", "Toolkit_Menu.bas", _
                     "Toolkit_RibbonQAT.bas", "Toolkit_Shortcuts.bas", "Toolkit_Sync.bas")
    
    Set compsContainer = Nothing
    On Error Resume Next
    Set compsContainer = w.NormalTemplate.VBProject.VBComponents
    If compsContainer Is Nothing And w.Documents.Count > 0 Then
        Set compsContainer = w.ActiveDocument.VBProject.VBComponents
    End If
    On Error GoTo 0
    
    If Not compsContainer Is Nothing Then
        For Each bFile In basFiles
            bName = Replace(bFile, ".bas", "")
            On Error Resume Next
            Set cObj = Nothing
            Set cObj = compsContainer.Item(bName)
            If Not cObj Is Nothing Then compsContainer.Remove cObj
            Err.Clear
            compsContainer.Import fso.BuildPath(appDir, bFile)
            On Error GoTo 0
        Next
        On Error Resume Next
        w.NormalTemplate.Save
        On Error GoTo 0
    End If
End If

On Error Resume Next
w.Run "Toolkit_Sync.SyncSelections", u, m, s, r
If Err.Number <> 0 Then
    MsgBox "Direct install failed." & vbCrLf & _
           "Please check 'Trust access to the VBA project object model' in Word Options -> Trust Center -> Macro Settings." & vbCrLf & _
           "Error: " & Err.Description, vbCritical, "Word Toolkit"
End If
On Error GoTo 0

WScript.Quit