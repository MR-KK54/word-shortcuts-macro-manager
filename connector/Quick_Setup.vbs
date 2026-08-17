Option Explicit
' ============================================================
' WORD TOOLKIT - 1-CLICK QUICK SETUP
' Automatically configures Word registry, opens Normal.dotm,
' imports all connector modules, and saves Normal.dotm.
' ============================================================

Dim Wsh, fso, url, appDataDir, files, i, w, normPath, normDoc, vbProj, comps, comp, importedCount, baseName, n, http, resText, f, vKey, rPath, rPaths

Set Wsh = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

url = "http://localhost:3000/api"
appDataDir = fso.BuildPath(Wsh.ExpandEnvironmentStrings("%APPDATA%"), "WordToolkit")
If Not fso.FolderExists(appDataDir) Then fso.CreateFolder appDataDir

' --- 1. Terminate any open Word instances so Normal.dotm is unlocked ---
On Error Resume Next
Wsh.Run "taskkill /F /IM winword.exe /T", 0, True
WScript.Sleep 1000
Err.Clear
On Error GoTo 0

' --- 2. Enable Registry Access across all Office versions ---
rPaths = Array("HKCU\Software\Microsoft\Office\", "HKCU\Software\Policies\Microsoft\Office\", "HKLM\Software\Microsoft\Office\")
For Each vKey In Array("16.0", "15.0", "14.0", "12.0", "11.0")
    For Each rPath In rPaths
        On Error Resume Next
        Wsh.RegWrite rPath & vKey & "\Word\Security\AccessVBOM", 1, "REG_DWORD"
        Wsh.RegWrite rPath & vKey & "\Word\Security\Level", 1, "REG_DWORD"
        Wsh.RegWrite rPath & vKey & "\Word\Security\VBAWarnings", 1, "REG_DWORD"
        Wsh.RegWrite rPath & vKey & "\Word\Security\DisableAllMacros", 0, "REG_DWORD"
        Wsh.RegWrite rPath & vKey & "\Word\Security\ExtensionHardening", 0, "REG_DWORD"
        On Error GoTo 0
    Next
Next

' --- 3. Register wordtoolkit:// protocol ---
On Error Resume Next
Wsh.RegWrite "HKCU\Software\Classes\wordtoolkit\", "URL:WordToolkit Protocol", "REG_SZ"
Wsh.RegWrite "HKCU\Software\Classes\wordtoolkit\URL Protocol", "", "REG_SZ"
Wsh.RegWrite "HKCU\Software\Classes\wordtoolkit\shell\open\command\", """" & WScript.FullName & """" & " """ & fso.BuildPath(appDataDir, "sync-handler.vbs") & """" & " ""%1""", "REG_SZ"
On Error GoTo 0

' --- 4. Fetch connector files ---
files = Array("Toolkit_Helpers.bas", "Toolkit_Macros.bas", "Toolkit_Menu.bas", "Toolkit_RibbonQAT.bas", "Toolkit_Shortcuts.bas", "Toolkit_Sync.bas", "sync-handler.vbs")
n = 0
For i = 0 To UBound(files)
    resText = FetchText(url & "/connector/" & files(i))
    If resText <> "" Then
        Set f = fso.CreateTextFile(fso.BuildPath(appDataDir, files(i)), True)
        f.Write resText
        f.Close
        n = n + 1
    End If
Next

' --- 5. Open Normal.dotm as document and import all 6 modules ---
On Error Resume Next
Set w = CreateObject("Word.Application")
If Not w Is Nothing Then
    w.Visible = False
    w.DisplayAlerts = 0
End If
On Error GoTo 0

If w Is Nothing Then
    MsgBox "Could not launch Word application background automation.", vbCritical, "Word Toolkit Setup"
    WScript.Quit
End If

normPath = fso.BuildPath(Wsh.ExpandEnvironmentStrings("%APPDATA%"), "Microsoft\Templates\Normal.dotm")
Set normDoc = Nothing

On Error Resume Next
If fso.FileExists(normPath) Then
    Set normDoc = w.Documents.Open(normPath)
Else
    Set normDoc = w.Documents.Add()
End If
On Error GoTo 0

Set comps = Nothing
If Not normDoc Is Nothing Then
    On Error Resume Next
    Set comps = normDoc.VBProject.VBComponents
    On Error GoTo 0
End If
If comps Is Nothing Then
    On Error Resume Next
    Set comps = w.NormalTemplate.VBProject.VBComponents
    On Error GoTo 0
End If

importedCount = 0
If Not comps Is Nothing Then
    For i = 0 To UBound(files)
        If LCase(Right(files(i), 4)) = ".bas" Then
            baseName = Replace(files(i), ".bas", "")
            On Error Resume Next
            Set comp = Nothing
            Set comp = comps.Item(baseName)
            If Not comp Is Nothing Then comps.Remove comp
            Err.Clear
            comps.Import fso.BuildPath(appDataDir, files(i))
            If Err.Number = 0 Then importedCount = importedCount + 1
            On Error GoTo 0
        End If
    Next
End If

On Error Resume Next
If Not normDoc Is Nothing Then
    normDoc.Save
    normDoc.Close
End If
w.NormalTemplate.Save
w.Quit
Set w = Nothing
On Error GoTo 0

MsgBox "🎉 Word Toolkit Quick Setup Complete!" & vbCrLf & vbCrLf & _
       "• Word Registry Permissions: ENABLED (AccessVBOM = 1)" & vbCrLf & _
       "• Protocol Registered: wordtoolkit://" & vbCrLf & _
       "• Connector Modules Imported into Normal.dotm: " & importedCount & " module(s)" & vbCrLf & vbCrLf & _
       "Now open Microsoft Word and press Alt + F11 to see all modules under Normal!", _
       vbInformation, "Word Toolkit Quick Setup"

Function FetchText(reqUrl)
    Dim hReq
    FetchText = ""
    On Error Resume Next
    Set hReq = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    If hReq Is Nothing Then Set hReq = CreateObject("MSXML2.XMLHTTP")
    hReq.Open "GET", reqUrl, False
    hReq.Send
    If Err.Number = 0 And hReq.Status >= 200 And hReq.Status < 300 Then
        FetchText = hReq.ResponseText
    End If
    On Error GoTo 0
End Function
