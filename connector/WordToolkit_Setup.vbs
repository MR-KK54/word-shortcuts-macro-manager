Option Explicit
' ============================================================
'  WORD TOOLKIT - ONE-TIME SETUP (run once per PC)
'  What it does:
'   1. Downloads the connector modules from your server
'   2. Registers the "wordtoolkit://" link so the web tool can
'      install macros/shortcuts/ribbon into Word with one click
'   3. Creates the Windows handler that receives those links
'   4. Imports the connector modules into Word (Normal project)
'  REQUIREMENT: File > Options > Trust Center > Macro Settings
'   -> enable "Trust access to the VBA project object model".
'  Close ALL other Word windows first, then double-click this file.
' ============================================================

Dim Wsh, fso, url, appDataDir, i, n, f, http, w, vbProj, comp, compName
Dim files, baseName, cmd, handlerPath

Set Wsh = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

url = InputBox("Enter your Word Toolkit server URL:" & vbCrLf & _
               "(e.g. https://your-app.onrender.com/api  or  http://localhost:3000/api)", _
               "Word Toolkit - One-Time Setup", "http://localhost:3000/api")
If url = "" Then WScript.Quit
If Right(url, 1) = "/" Then url = Left(url, Len(url) - 1)
If LCase(Right(url, 4)) <> "/api" Then url = url & "/api"

' --- 1. Download connector modules ---
appDataDir = fso.BuildPath(Wsh.ExpandEnvironmentStrings("%APPDATA%"), "WordToolkit")
If Not fso.FolderExists(appDataDir) Then fso.CreateFolder appDataDir

files = Array("Toolkit_Helpers.bas", "Toolkit_Macros.bas", "Toolkit_Menu.bas", _
              "Toolkit_RibbonQAT.bas", "Toolkit_Shortcuts.bas", "Toolkit_Sync.bas")
n = 0
For i = 0 To UBound(files)
    Set http = CreateObject("MSXML2.XMLHTTP")
    http.Open "GET", url & "/connector/" & files(i), False
    http.Send
    If http.Status >= 200 And http.Status < 300 Then
        Set f = fso.CreateTextFile(fso.BuildPath(appDataDir, files(i)), True, True)
        f.Write http.ResponseText
        f.Close
        n = n + 1
    Else
        MsgBox "Could not download " & files(i) & " (HTTP " & http.Status & ")." & vbCrLf & _
               "Check the server URL.", vbExclamation, "Word Toolkit Setup"
    End If
Next
If n = 0 Then
    MsgBox "No connector files could be downloaded. Setup aborted.", vbCritical, "Word Toolkit Setup"
    WScript.Quit
End If

' --- 2. Register the wordtoolkit:// protocol link ---
cmd = """" & Wsh.ExpandEnvironmentStrings("%WINDIR%") & "\System32\wscript.exe"" """ & _
      appDataDir & "\sync-handler.vbs"" ""%1"""
On Error Resume Next
Wsh.RegWrite "HKCU\Software\Classes\wordtoolkit\shell\open\command\", cmd, "REG_SZ"
If Err.Number <> 0 Then
    MsgBox "Could not register the wordtoolkit:// link." & vbCrLf & "Error: " & Err.Description, vbCritical, "Word Toolkit Setup"
    WScript.Quit
End If
On Error GoTo 0

' --- 3. Write the link handler script ---
handlerPath = fso.BuildPath(appDataDir, "sync-handler.vbs")
Set f = fso.CreateTextFile(handlerPath, True, True)
f.Write HandlerSource()
f.Close

' --- 4. Enable Word import access automatically (trust + enable all macros) ---
EnableWordAccess Wsh

' --- 5. Import connector modules into Word ---
On Error Resume Next
Set w = GetObject(, "Word.Application")
If w Is Nothing Then
    Set w = CreateObject("Word.Application")
    w.Visible = True
End If
If Err.Number <> 0 Then
    MsgBox "Could not start Microsoft Word." & vbCrLf & "Error: " & Err.Description, vbCritical, "Word Toolkit Setup"
    WScript.Quit
End If
On Error GoTo 0

On Error Resume Next
Set vbProj = w.VBE.ActiveVBProject
For i = 0 To UBound(files)
    baseName = Replace(files(i), ".bas", "")
    Set comp = Nothing
    Set comp = vbProj.VBComponents(baseName)
    If Not comp Is Nothing Then vbProj.VBComponents.Remove comp
    vbProj.VBComponents.Import fso.BuildPath(appDataDir, files(i))
    comp = Nothing
    If Err.Number <> 0 Then
        MsgBox "Importing " & files(i) & " failed." & vbCrLf & _
               "Enable 'Trust access to the VBA project object model' and try again." & vbCrLf & _
               "Error: " & Err.Description, vbCritical, "Word Toolkit Setup"
        WScript.Quit
    End If
Next
On Error GoTo 0

' --- 5. Tell Word to remember the modules ---
On Error Resume Next
w.NormalTemplate.Saved = False
On Error GoTo 0

MsgBox "Word Toolkit setup complete!" & vbCrLf & vbCrLf & _
       n & " connector module(s) downloaded from " & url & vbCrLf & _
       "Modules imported into Word (remember to save Normal.dotm when closing Word)." & vbCrLf & _
       "Word import access enabled automatically (Trust access + enable all macros)." & vbCrLf & _
       "The wordtoolkit:// link is now registered." & vbCrLf & vbCrLf & _
       "You can now click 'Export to Word' on the web tool to install directly.", _
       vbInformation, "Word Toolkit Setup"

' ============================================================
' Enables Word's import permissions for the toolkit:
'  - AccessVBOM : "Trust access to the VBA project object model"
'  - Level      : enable all macros
'  - VBAWarnings: no warnings for unsigned macros
' Applied for every installed Office version key. A running
' Word reads these at startup, so restart Word to apply.
' ============================================================
Sub EnableWordAccess(Wsh)
    Dim versions, v, secKey
    versions = Array("16.0", "15.0", "14.0", "12.0", "11.0")
    For Each v In versions
        secKey = "HKCU\Software\Microsoft\Office\" & v & "\Word\Security\"
        On Error Resume Next
        Wsh.RegWrite secKey & "AccessVBOM", 1, "REG_DWORD"
        Wsh.RegWrite secKey & "Level", 1, "REG_DWORD"
        Wsh.RegWrite secKey & "VBAWarnings", 1, "REG_DWORD"
        On Error GoTo 0
    Next
End Sub

' ============================================================
' Returns the source of the sync-handler.vbs written above.
' (Kept as a string so this file stays fully self-contained.)
' Every direct install re-asserts the Word access settings.
' ============================================================
Function HandlerSource()
    Dim h
    h = "Option Explicit" & vbCrLf
    h = h & "'' Word Toolkit - receives wordtoolkit://sync links from the web tool" & vbCrLf
    h = h & "'' and tells the running Microsoft Word to install the selection." & vbCrLf
    h = h & "Dim argLine, parts, kv, i, dict, p, k, u, m, s, r, w, Wsh, v, secKey" & vbCrLf
    h = h & "argLine = """ & vbCrLf
    h = h & "If WScript.Arguments.Count > 0 Then argLine = WScript.Arguments(0)" & vbCrLf
    h = h & "Set dict = CreateObject(""Scripting.Dictionary"")" & vbCrLf
    h = h & "parts = Split(argLine, ""&"")" & vbCrLf
    h = h & "For i = 0 To UBound(parts)" & vbCrLf
    h = h & "    kv = Split(parts(i), ""="")" & vbCrLf
    h = h & "    If UBound(kv) >= 1 Then" & vbCrLf
    h = h & "        p = kv(0)" & vbCrLf
    h = h & "        k = Mid(p, InStrRev(p, ""?"") + 1)" & vbCrLf
    h = h & "        dict(k) = Unescape(kv(1))" & vbCrLf
    h = h & "    End If" & vbCrLf
    h = h & "Next" & vbCrLf
    h = h & "'' Enable Word import access automatically every time" & vbCrLf
    h = h & "Set Wsh = CreateObject(""WScript.Shell"")" & vbCrLf
    h = h & "For Each v In Array(""16.0"", ""15.0"", ""14.0"", ""12.0"", ""11.0"")" & vbCrLf
    h = h & "    secKey = ""HKCU\Software\Microsoft\Office\"" & v & ""\Word\Security\""" & vbCrLf
    h = h & "    On Error Resume Next" & vbCrLf
    h = h & "    Wsh.RegWrite secKey & ""AccessVBOM"", 1, ""REG_DWORD""" & vbCrLf
    h = h & "    Wsh.RegWrite secKey & ""Level"", 1, ""REG_DWORD""" & vbCrLf
    h = h & "    Wsh.RegWrite secKey & ""VBAWarnings"", 1, ""REG_DWORD""" & vbCrLf
    h = h & "    On Error GoTo 0" & vbCrLf
    h = h & "Next" & vbCrLf
    h = h & "'' Special action: just enable access (web tool button)" & vbCrLf
    h = h & "If dict.Exists(""act"") And dict(""act"") = ""enable"" Then" & vbCrLf
    h = h & "    MsgBox ""Word import access enabled (Trust access + all macros)."" & vbCrLf & ""Restart Word once, then install again."" & vbCrLf & ""Nothing was imported."", vbInformation, ""Word Toolkit""" & vbCrLf
    h = h & "    WScript.Quit" & vbCrLf
    h = h & "End If" & vbCrLf
    h = h & "u = dict(""u""): m = dict(""m""): s = dict(""s""): r = dict(""r"")" & vbCrLf
    h = h & "On Error Resume Next" & vbCrLf
    h = h & "Set w = GetObject(, ""Word.Application"")" & vbCrLf
    h = h & "If w Is Nothing Then" & vbCrLf
    h = h & "    Set w = CreateObject(""Word.Application"")" & vbCrLf
    h = h & "    w.Visible = True" & vbCrLf
    h = h & "End If" & vbCrLf
    h = h & "If Err.Number <> 0 Then" & vbCrLf
    h = h & "    MsgBox ""Could not start Word: "" & Err.Description, vbCritical, ""Word Toolkit""" & vbCrLf
    h = h & "    WScript.Quit" & vbCrLf
    h = h & "End If" & vbCrLf
    h = h & "On Error GoTo 0" & vbCrLf
    h = h & "On Error Resume Next" & vbCrLf
    h = h & "w.Run ""Toolkit_Sync.SyncSelections"", u, m, s, r" & vbCrLf
    h = h & "If Err.Number <> 0 Then" & vbCrLf
    h = h & "    MsgBox ""Direct install failed. Run the setup again."" & vbCrLf & ""Error: "" & Err.Description, vbCritical, ""Word Toolkit""" & vbCrLf
    h = h & "End If" & vbCrLf
    h = h & "'' The SyncSelections macro shows its own summary box." & vbCrLf
    h = h & "WScript.Quit" & vbCrLf
    HandlerSource = h
End Function