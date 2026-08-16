Option Explicit
' ============================================================
'  WORD TOOLKIT - ONE-TIME SETUP (run once per PC)
'  What it does:
'   1. Downloads the connector modules + link handler from your server
'   2. Registers the "wordtoolkit://" link so the web tool can
'      install macros/shortcuts/ribbon into Word with one click
'   3. Enables Word's import access automatically
'      (Trust access to VBA project object model + all macros)
'   4. Imports the connector modules into Word (Normal project)
'  REQUIREMENT: close ALL other Word windows first, then
'  double-click this file.
' ============================================================

Dim Wsh, fso, url, appDataDir, i, n, f, http, w, vbProj, comp, compName
Dim files, baseName, cmd, handlerPath, c

Set Wsh = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

url = InputBox("Enter your Word Toolkit server URL:" & vbCrLf & _
               "(e.g. https://your-app.onrender.com/api  or  http://localhost:3000/api)", _
               "Word Toolkit - One-Time Setup", "http://localhost:3000/api")
If url = "" Then WScript.Quit
If Right(url, 1) = "/" Then url = Left(url, Len(url) - 1)
If LCase(Right(url, 4)) <> "/api" Then url = url & "/api"

' --- 1. Download connector modules + handler ---
appDataDir = fso.BuildPath(Wsh.ExpandEnvironmentStrings("%APPDATA%"), "WordToolkit")
If Not fso.FolderExists(appDataDir) Then fso.CreateFolder appDataDir

files = Array("Toolkit_Helpers.bas", "Toolkit_Macros.bas", "Toolkit_Menu.bas", _
              "Toolkit_RibbonQAT.bas", "Toolkit_Shortcuts.bas", "Toolkit_Sync.bas", _
              "sync-handler.vbs")
n = 0
For i = 0 To UBound(files)
    On Error Resume Next
    Set http = CreateObject("MSXML2.XMLHTTP")
    http.Open "GET", url & "/connector/" & files(i), False
    http.Send
    If Err.Number = 0 And http.Status >= 200 And http.Status < 300 Then
        Set f = fso.CreateTextFile(fso.BuildPath(appDataDir, files(i)), True, True)
        f.Write http.ResponseText
        f.Close
        n = n + 1
    Else
        MsgBox "Could not download " & files(i) & " (HTTP " & http.Status & ")." & vbCrLf & _
               "Check the server URL or your internet connection." & vbCrLf & _
               "Error: " & Err.Description, vbExclamation, "Word Toolkit Setup"
    End If
    On Error GoTo 0
Next
If n = 0 Then
    MsgBox "No connector files could be downloaded. Setup aborted.", vbCritical, "Word Toolkit Setup"
    WScript.Quit
End If

' --- 2. Register the wordtoolkit:// protocol link (points at the handler) ---
cmd = """" & Wsh.ExpandEnvironmentStrings("%WINDIR%") & "\System32\wscript.exe"" """ & _
      appDataDir & "\sync-handler.vbs"" ""%1"""
On Error Resume Next
Wsh.RegWrite "HKCU\Software\Classes\wordtoolkit\shell\open\command\", cmd, "REG_SZ"
If Err.Number <> 0 Then
    MsgBox "Could not register the wordtoolkit:// link." & vbCrLf & "Error: " & Err.Description, vbCritical, "Word Toolkit Setup"
    WScript.Quit
End If
On Error GoTo 0

' --- 3. Enable Word import access automatically (Trust access + all macros) ---
EnableWordAccess Wsh

' --- 4. Import connector modules into Word (handler is NOT imported) ---
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
    If LCase(Right(files(i), 4)) = ".bas" Then
        baseName = Replace(files(i), ".bas", "")
        Set comp = Nothing
        Set comp = vbProj.VBComponents(baseName)
        If Not comp Is Nothing Then vbProj.VBComponents.Remove comp
        vbProj.VBComponents.Import fso.BuildPath(appDataDir, files(i))
        comp = Nothing
        If Err.Number <> 0 Then
            MsgBox "Importing " & files(i) & " failed." & vbCrLf & _
                   "Word import access was enabled automatically - restart Word and run this setup again." & vbCrLf & _
                   "Error: " & Err.Description, vbCritical, "Word Toolkit Setup"
            WScript.Quit
        End If
    End If
Next
On Error GoTo 0

' --- 5. Tell Word to remember the modules ---
On Error Resume Next
w.NormalTemplate.Saved = False
On Error GoTo 0

MsgBox "Word Toolkit setup complete!" & vbCrLf & vbCrLf & _
       n & " file(s) downloaded from " & url & vbCrLf & _
       "Connector modules imported into Word (remember to save Normal.dotm when closing Word)." & vbCrLf & _
       "Word import access enabled automatically (Trust access + enable all macros)." & vbCrLf & _
       "The wordtoolkit:// link is registered - it stays up to date by itself." & vbCrLf & vbCrLf & _
       "Now click 'Export to Word' or 'Enable Word Import Access' on the web tool.", _
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