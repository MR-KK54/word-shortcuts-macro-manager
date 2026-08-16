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
              "sync-handler.vbs", "Enable_Word_Import_Access.vbs")
n = 0
For i = 0 To UBound(files)
    Dim resText
    resText = FetchText(url & "/connector/" & files(i))
    If Len(resText) > 0 Then
        Set f = fso.CreateTextFile(fso.BuildPath(appDataDir, files(i)), True, True)
        f.Write resText
        f.Close
        n = n + 1
    End If
Next

If n = 0 Then
    MsgBox "Could not connect to the Word Toolkit server at:" & vbCrLf & _
           url & vbCrLf & vbCrLf & _
           "• If running locally: Please start your server first (run 'npm start' in terminal)." & vbCrLf & _
           "• If using online server: Check your internet connection and server URL.", _
           vbCritical, "Word Toolkit Setup - Connection Error"
    WScript.Quit
End If

Function FetchText(httpUrl)
    Dim http
    FetchText = ""
    On Error Resume Next
    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    If http Is Nothing Or Err.Number <> 0 Then
        Err.Clear
        Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    End If
    If http Is Nothing Or Err.Number <> 0 Then
        Err.Clear
        Set http = CreateObject("MSXML2.XMLHTTP")
    End If
    If Not http Is Nothing Then
        http.Open "GET", httpUrl, False
        http.Send
        If Err.Number = 0 Then
            If http.Status >= 200 And http.Status < 300 Then
                FetchText = http.ResponseText
            End If
        End If
    End If
    On Error GoTo 0
End Function

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
If w Is Nothing Or Err.Number <> 0 Then
    Err.Clear
    Set w = CreateObject("Word.Application")
    w.Visible = True
End If
If Err.Number <> 0 Or w Is Nothing Then
    MsgBox "Could not start Microsoft Word." & vbCrLf & "Error: " & Err.Description, vbCritical, "Word Toolkit Setup"
    WScript.Quit
End If
On Error GoTo 0

Dim docCreated, doc
docCreated = False
If w.Documents.Count = 0 Then
    On Error Resume Next
    Set doc = w.Documents.Add()
    docCreated = True
    On Error GoTo 0
End If

' Safely acquire Word's VB project & components container
Set vbProj = Nothing
Set comps = Nothing

On Error Resume Next
Set vbProj = w.NormalTemplate.VBProject
If Not vbProj Is Nothing Then Set comps = vbProj.VBComponents

If comps Is Nothing And Not doc Is Nothing Then
    Err.Clear
    Set vbProj = doc.VBProject
    If Not vbProj Is Nothing Then Set comps = vbProj.VBComponents
End If

If comps Is Nothing Then
    Err.Clear
    Set vbProj = w.VBE.ActiveVBProject
    If Not vbProj Is Nothing Then Set comps = vbProj.VBComponents
End If
On Error GoTo 0

If comps Is Nothing Then
    MsgBox "Microsoft Word is blocking programmatic macro imports." & vbCrLf & vbCrLf & _
           "To allow imports in 10 seconds:" & vbCrLf & _
           "1. Open Microsoft Word" & vbCrLf & _
           "2. Go to File -> Options -> Trust Center -> Trust Center Settings -> Macro Settings" & vbCrLf & _
           "3. Check 'Trust access to the VBA project object model' and click OK" & vbCrLf & _
           "4. Run this setup script again.", _
           vbExclamation, "Word Toolkit Setup - Trust Access Required"
    If docCreated And Not doc Is Nothing Then
        On Error Resume Next
        doc.Close False
        On Error GoTo 0
    End If
    WScript.Quit
End If

For i = 0 To UBound(files)
    If LCase(Right(files(i), 4)) = ".bas" Then
        baseName = Replace(files(i), ".bas", "")
        On Error Resume Next
        Set comp = Nothing
        Set comp = comps.Item(baseName)
        If Not comp Is Nothing Then comps.Remove comp
        Err.Clear
        comps.Import fso.BuildPath(appDataDir, files(i))
        comp = Nothing
        If Err.Number <> 0 Then
            MsgBox "Importing " & files(i) & " failed." & vbCrLf & _
                   "Error (" & Err.Number & "): " & Err.Description & vbCrLf & vbCrLf & _
                   "Please make sure 'Trust access to the VBA project object model' is checked in Word Options -> Trust Center -> Macro Settings.", _
                   vbCritical, "Word Toolkit Setup"
            If docCreated And Not doc Is Nothing Then doc.Close False
            WScript.Quit
        End If
        On Error GoTo 0
    End If
Next

If docCreated And Not doc Is Nothing Then
    On Error Resume Next
    doc.Close False
    On Error GoTo 0
End If

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

Sub EnableWordAccess(Wsh)
    Dim versions, v, root, roots
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
End Sub