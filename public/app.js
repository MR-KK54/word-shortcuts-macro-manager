// Word Toolkit & Macro Manager - Frontend Client Logic

const $ = (sel, el = document) => el.querySelector(sel);
const $$ = (sel, el = document) => Array.from(el.querySelectorAll(sel));

// State Management
let apiBaseUrl = localStorage.getItem('wt_server_url') || '/api';
if (apiBaseUrl.endsWith('/')) apiBaseUrl = apiBaseUrl.slice(0, -1);

let isServerOnline = false;
let currentMacros = [];
let currentShortcuts = [];
let currentRibbon = [];

/* ---------- INIT & CONNECTOR FILES ---------- */
const CONNECTOR_FILES = {
'Toolkit_Helpers.bas': `Attribute VB_Name = "Toolkit_Helpers"
Option Explicit

' Helper function to prompt user for a folder path
Function GetFolder(promptText As String) As String
    Dim fd As Object
    On Error Resume Next
    Set fd = Application.FileDialog(4) ' 4 = msoFileDialogFolderPicker
    If Not fd Is Nothing Then
        fd.Title = promptText
        If fd.Show = -1 Then
            GetFolder = fd.SelectedItems(1)
            Exit Function
        End If
    End If
    GetFolder = ""
End Function
`,
'Toolkit_Macros.bas': `Attribute VB_Name = "Toolkit_Macros"
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
`,
'Toolkit_Menu.bas': `Attribute VB_Name = "Toolkit_Menu"
Option Explicit

Sub ShowToolkitMenu()
    Dim choice As String, msg As String
    msg = "WORD CUSTOMIZATION TOOLKIT" & vbCrLf & vbCrLf & _
          "  1 - Export Macros (To Files)" & vbCrLf & _
          "  2 - Import Macros (Replace Existing)" & vbCrLf & _
          "  3 - Export Keyboard Shortcuts (.CSV)" & vbCrLf & _
          "  4 - Import Keyboard Shortcuts (.CSV)" & vbCrLf & _
          "  5 - Export Ribbon / QAT Customizations" & vbCrLf & _
          "  6 - Import Ribbon / QAT Customizations" & vbCrLf & _
          "  7 - Direct Sync ALL From Server" & vbCrLf & vbCrLf & _
          "Enter an option number (1-7) or click Cancel:"
    choice = InputBox(msg, "Word Customization Toolkit")
    Select Case choice
        Case "1": ExportAllMacros
        Case "2": ImportAllMacros
        Case "3": ExportKeyboardShortcuts
        Case "4": ImportKeyboardShortcuts
        Case "5": ExportRibbonCustomizations
        Case "6": ImportRibbonCustomizations
        Case "7": SyncAllFromServer
        Case ""
        Case Else
            MsgBox "Please enter a number from 1 to 7.", vbExclamation, "Word Customization Toolkit"
    End Select
End Sub
`,
'Toolkit_RibbonQAT.bas': `Attribute VB_Name = "Toolkit_RibbonQAT"
Option Explicit

Private Function OfficeUIPath() As String
    OfficeUIPath = Environ$("LOCALAPPDATA") & "\Microsoft\Office\Word.officeUI"
End Function

Sub ExportRibbonCustomizations()
    Dim srcFile As String, destFolder As String, fso As Object
    srcFile = OfficeUIPath()
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(srcFile) Then
        MsgBox "No custom ribbon/QAT file was found at:" & vbCrLf & srcFile & vbCrLf & vbCrLf & "Use File > Options > Customize Ribbon > Import/Export in Word instead.", vbExclamation, "Export Ribbon/QAT"
        Exit Sub
    End If
    destFolder = GetFolder("Select a folder to export the ribbon/QAT customization to")
    If destFolder = "" Then Exit Sub
    On Error GoTo HandleErr
    fso.CopyFile srcFile, destFolder & "\Word.officeUI", True
    MsgBox "Ribbon and Quick Access Toolbar customizations exported to:" & vbCrLf & destFolder, vbInformation, "Export Ribbon/QAT"
    Exit Sub
HandleErr:
    MsgBox "Could not export the ribbon/QAT file." & vbCrLf & "Error: " & Err.Description, vbCritical, "Export Ribbon/QAT Error"
End Sub

Sub ImportRibbonCustomizations()
    Dim srcFile As Variant, destFile As String, fso As Object
    srcFile = Application.GetOpenFilename("Office UI Files (*.officeUI), *.officeUI")
    If srcFile = False Or srcFile = "False" Then Exit Sub
    destFile = OfficeUIPath()
    Set fso = CreateObject("Scripting.FileSystemObject")
    On Error GoTo HandleErr
    fso.CopyFile srcFile, destFile, True
    MsgBox "Ribbon/QAT customization imported. Close and reopen Word for changes to take effect.", vbInformation, "Import Ribbon/QAT"
    Exit Sub
HandleErr:
    MsgBox "Could not import the ribbon/QAT file. If Word reports it is in use, close Word and copy it manually to:" & vbCrLf & destFile & vbCrLf & vbCrLf & "Error: " & Err.Description, vbCritical, "Import Ribbon/QAT Error"
End Sub
`,
'Toolkit_Shortcuts.bas': `Attribute VB_Name = "Toolkit_Shortcuts"
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
`,
'Toolkit_Sync.bas': `Attribute VB_Name = "Toolkit_Sync"
Option Explicit

' ============================================================
'  DIRECT WORD SYNC - installs macros, keyboard shortcuts and
'  ribbon profiles from the server straight into Microsoft Word
'  (no file dialogs).
'  REQUIREMENT: Enable "Trust access to the VBA project object
'  model" (File > Options > Trust Center > Macro Settings).
'
'  ENTRY POINTS:
'   - SyncAllFromServer : menu option 7; uses the selections
'     baked into this module (web app "Export to Word" dialog).
'   - SyncSelections    : called directly by Windows when you
'     click "Install into Word" on the web page (wordtoolkit://
'     link registered by WordToolkit_Setup.vbs). Takes the
'     server URL and group choices as arguments - nothing baked in.
' ============================================================

' --- INSTALLER SELECTIONS (replaced by the web app) ---
' "" (empty) = sync ALL groups of that type.
Private Const SYNC_MACROS As Boolean = True
Private Const SYNC_SHORTCUTS As Boolean = True
Private Const SYNC_RIBBON As Boolean = True
Private Const DEFAULT_SERVER As String = "@@SERVER_URL@@"
Private Const SELECTED_MACRO_GROUP As String = "@@MACRO_GROUP@@"
Private Const SELECTED_SHORTCUT_GROUP As String = "@@SHORTCUT_GROUP@@"
Private Const SELECTED_RIBBON_GROUP As String = "@@RIBBON_GROUP@@"

' Ask once per session, remember the answer
Private Function ServerBaseUrl() As String
    Static sUrl As String
    Dim defUrl As String
    If sUrl = "" Then
        defUrl = DEFAULT_SERVER
        If defUrl = "" Or InStr(defUrl, "@@") > 0 Then defUrl = "http://localhost:3000/api"
        sUrl = InputBox("Enter the Word Toolkit server URL (e.g. https://your-app.onrender.com/api):", _
                        "Sync from Server", defUrl)
        If sUrl = "" Then Exit Function
        If Right(sUrl, 1) = "/" Then sUrl = Left(sUrl, Len(sUrl) - 1)
    End If
    ServerBaseUrl = sUrl
End Function

' Builds the sync URL for one section; grp="" falls back to the
' baked-in constant of that section, "@@" (unset) means all groups.
Private Function SectionUrl(baseUrl As String, endpoint As String, grp As String, defGrp As String) As String
    Dim g As String
    g = grp
    If g = "" Then g = defGrp
    SectionUrl = baseUrl & "/" & endpoint
    If g <> "" And InStr(g, "@@") = 0 Then
        SectionUrl = SectionUrl & "?group=" & Replace(g, " ", "%20")
    End If
End Function

' Simple HTTP GET returning the response body as text
Private Function HttpGetText(url As String) As String
    Dim http As Object
    On Error GoTo HttpErr
    Set http = CreateObject("MSXML2.XMLHTTP")
    http.Open "GET", url, False
    http.Send
    If http.Status >= 200 And http.Status < 300 Then
        HttpGetText = http.ResponseText
    Else
        HttpGetText = ""
        MsgBox "Server returned status " & http.Status, vbExclamation, "Sync from Server"
    End If
    Exit Function
HttpErr:
    MsgBox "Could not reach the server." & vbCrLf & "Check the URL and that you are online." & vbCrLf & "Error: " & Err.Description, vbCritical, "Sync from Server"
    HttpGetText = ""
End Function

' --- DIRECT INSTALL ENTRY (triggered by the web page via wordtoolkit://) ---
Public Sub SyncSelections(baseUrl As String, macGroup As String, scGroup As String, rbGroup As String)
    Dim url As String, macMsg As String, scMsg As String, rbMsg As String
    url = baseUrl
    If url = "" Then url = ServerBaseUrl()
    If url = "" Then Exit Sub
    If Right(url, 1) = "/" Then url = Left(url, Len(url) - 1)

    macMsg = SyncMacrosFromServer(url, macGroup)
    scMsg = SyncShortcutsFromServer(url, scGroup)
    rbMsg = SyncRibbonFromServer(url, rbGroup)

    MsgBox "Word Toolkit - Direct Install complete!" & vbCrLf & vbCrLf & _
           macMsg & vbCrLf & scMsg & vbCrLf & rbMsg & vbCrLf & vbCrLf & _
           "Save Normal.dotm when you close Word so the changes persist.", _
           vbInformation, "Word Toolkit Direct Install"
End Sub

' --- MENU ENTRY (menu option 7): uses the baked-in selections ---
Sub SyncAllFromServer()
    Dim baseUrl As String, macMsg As String, scMsg As String, rbMsg As String
    baseUrl = ServerBaseUrl()
    If baseUrl = "" Then Exit Sub

    macMsg = "Macros: skipped."
    If SYNC_MACROS Then macMsg = SyncMacrosFromServer(baseUrl)
    scMsg = "Shortcuts: skipped."
    If SYNC_SHORTCUTS Then scMsg = SyncShortcutsFromServer(baseUrl)
    rbMsg = "Ribbon: skipped."
    If SYNC_RIBBON Then rbMsg = SyncRibbonFromServer(baseUrl)

    MsgBox "Direct sync complete!" & vbCrLf & vbCrLf & _
           macMsg & vbCrLf & scMsg & vbCrLf & rbMsg & vbCrLf & vbCrLf & _
           "Remember to save Normal.dotm or your active template.", _
           vbInformation, "Word Toolkit Sync"
End Sub

' --- MACROS ---
Public Function SyncMacrosFromServer(baseUrl As String, Optional grp As String = "") As String
    Dim bundle As String, vbProj As Object, count As Integer, replacedCount As Integer
    Dim groupName As String, compName As String, compType As String, code As String
    Dim lines() As String, i As Long, inMacro As Boolean
    Dim fso As Object, tmpFile As String, fnum As Integer, existingComp As Object
    Dim ln As String

    On Error GoTo HandleErr

    bundle = HttpGetText(SectionUrl(baseUrl, "sync/macros", grp, SELECTED_MACRO_GROUP))
    If bundle = "" Then
        SyncMacrosFromServer = "Macros: not synced (no data or server unreachable)."
        Exit Function
    End If

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set vbProj = Application.VBE.ActiveVBProject

    lines = Split(bundle, vbLf)
    inMacro = False
    code = ""
    groupName = ""
    compName = ""
    compType = "bas"

    For i = 0 To UBound(lines)
        ln = Trim(lines(i))
        If ln = "#WTMACRO-END#" Then
            ' close current macro and install it
            If inMacro And compName <> "" And Len(code) > 0 Then
                tmpFile = Environ$("TEMP") & "\\wtk_" & compName & "." & compType
                fnum = FreeFile
                Open tmpFile For Output As #fnum
                Print #fnum, code
                Close #fnum

                ' Replace existing module with same name
                On Error Resume Next
                Set existingComp = vbProj.VBComponents(compName)
                If Not existingComp Is Nothing Then
                    vbProj.VBComponents.Remove existingComp
                    replacedCount = replacedCount + 1
                End If
                On Error GoTo HandleErr

                vbProj.VBComponents.Import tmpFile
                fso.DeleteFile tmpFile, True
                count = count + 1
            End If
            inMacro = False
            code = ""
            compName = ""
            groupName = ""
            compType = "bas"
        ElseIf Left(ln, 7) = "@group=" Then
            groupName = Mid(ln, 8)
        ElseIf Left(ln, 6) = "@name=" Then
            compName = Mid(ln, 7)
            inMacro = True
        ElseIf Left(ln, 6) = "@type=" Then
            compType = LCase(Mid(ln, 7))
        ElseIf inMacro Then
            ' code line
            code = code & lines(i) & vbCrLf
        End If
    Next i

    SyncMacrosFromServer = "Macros: " & count & " installed (" & replacedCount & " replaced)."
    Exit Function

HandleErr:
    MsgBox "Could not install macros. Enable 'Trust access to the VBA project object model'." & vbCrLf & _
           "Error: " & Err.Description, vbCritical, "Sync Macros Error"
    SyncMacrosFromServer = "Macros: failed with error."
End Function

' --- SHORTCUTS ---
Public Function SyncShortcutsFromServer(baseUrl As String, Optional grp As String = "") As String
    Dim csvData As String, lines() As String, i As Long
    Dim parts() As String, count As Integer, failCount As Integer
    Dim ln As String, setCount As Integer

    On Error GoTo HandleErr

    csvData = HttpGetText(SectionUrl(baseUrl, "sync/shortcuts", grp, SELECTED_SHORTCUT_GROUP))
    If csvData = "" Then
        SyncShortcutsFromServer = "Shortcuts: not synced (no data or server unreachable)."
        Exit Function
    End If

    Application.CustomizationContext = NormalTemplate
    lines = Split(csvData, vbLf)

    For i = 0 To UBound(lines)
        ln = Trim(lines(i))
        If ln = "" Then
            ' skip
        ElseIf Left(ln, 5) = "#SET:" Then
            setCount = setCount + 1
        ElseIf Left(ln, 11) = "KeyCategory" Then
            ' header - skip
        Else
            parts = Split(ln, ",")
            If UBound(parts) >= 3 Then
                On Error Resume Next
                Err.Clear
                Application.KeyBindings.Add KeyCategory:=CLng(parts(0)), Command:=parts(1), _
                    KeyCode:=CLng(parts(2)), KeyCode2:=IIf(parts(3) = "0" Or parts(3) = "", 0, CLng(parts(3)))
                If Err.Number = 0 Then count = count + 1 Else failCount = failCount + 1
                On Error GoTo HandleErr
            End If
        End If
    Next i

    SyncShortcutsFromServer = "Shortcuts: " & count & " applied from " & setCount & " set(s)" & _
                              IIf(failCount > 0, " (" & failCount & " failed).", ".")
    Exit Function

HandleErr:
    MsgBox "Could not apply keyboard shortcuts." & vbCrLf & "Error: " & Err.Description, vbCritical, "Sync Shortcuts Error"
    SyncShortcutsFromServer = "Shortcuts: failed with error."
End Function

' --- RIBBON ---
' Downloads the selected .officeUI profile(s) and writes the file to
' Word's AppData folder. Word must be restarted to apply the ribbon.
Public Function SyncRibbonFromServer(baseUrl As String, Optional grp As String = "") As String
    Dim bundle As String, lines() As String, i As Long
    Dim ln As String, fso As Object, folderPath As String, filePath As String
    Dim fnum As Integer, xml As String, written As Integer, inRibbon As Boolean

    On Error GoTo HandleErr

    bundle = HttpGetText(SectionUrl(baseUrl, "sync/ribbon", grp, SELECTED_RIBBON_GROUP))
    If bundle = "" Then
        SyncRibbonFromServer = "Ribbon: not synced (no data or server unreachable)."
        Exit Function
    End If

    Set fso = CreateObject("Scripting.FileSystemObject")
    folderPath = Environ$("APPDATA") & "\\Microsoft\\Office"
    If Not fso.FolderExists(folderPath) Then fso.CreateFolder folderPath
    filePath = folderPath & "\\Word.officeUI"

    lines = Split(bundle, vbLf)
    inRibbon = False
    xml = ""

    For i = 0 To UBound(lines)
        ln = Trim(lines(i))
        If Left(ln, 8) = "#RIBBON:" Then
            ' A new profile begins: flush the previous one first
            If inRibbon And Len(xml) > 0 Then
                fnum = FreeFile
                Open filePath For Output As #fnum
                Print #fnum, xml
                Close #fnum
                written = written + 1
            End If
            inRibbon = True
            xml = ""
        ElseIf ln = "#WTRIBBON-END#" Then
            If inRibbon And Len(xml) > 0 Then
                fnum = FreeFile
                Open filePath For Output As #fnum
                Print #fnum, xml
                Close #fnum
                written = written + 1
            End If
            inRibbon = False
            xml = ""
        ElseIf inRibbon Then
            xml = xml & lines(i) & vbCrLf
        End If
    Next i

    If written = 0 Then
        SyncRibbonFromServer = "Ribbon: no profiles found."
    Else
        SyncRibbonFromServer = "Ribbon: " & written & " profile(s) written to Word.officeUI. " & _
                               "Restart Word to apply."
    End If
    Exit Function

HandleErr:
    MsgBox "Could not install the ribbon profile." & vbCrLf & _
           "Error: " & Err.Description, vbCritical, "Sync Ribbon Error"
    SyncRibbonFromServer = "Ribbon: failed with error."
End Function`,
'WordToolkit_Setup.vbs': `Option Explicit
' ============================================================
'  WORD TOOLKIT - ONE-TIME SETUP (run once per PC)
' ============================================================

Dim Wsh, fso, url, appDataDir, i, n, f, w, vbProj, comp, compName
Dim files, baseName, cmd, handlerPath, c, comps, docCreated, doc, importedCount, resText, normDoc, normProj

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
    Dim hReq
    FetchText = ""
    On Error Resume Next
    Set hReq = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    If hReq Is Nothing Or Err.Number <> 0 Then
        Err.Clear
        Set hReq = CreateObject("WinHttp.WinHttpRequest.5.1")
    End If
    If hReq Is Nothing Or Err.Number <> 0 Then
        Err.Clear
        Set hReq = CreateObject("MSXML2.XMLHTTP")
    End If
    If Not hReq Is Nothing Then
        hReq.Open "GET", httpUrl, False
        hReq.Send
        If Err.Number = 0 Then
            If hReq.Status >= 200 And hReq.Status < 300 Then
                FetchText = hReq.ResponseText
            End If
        End If
    End If
    On Error GoTo 0
End Function

' --- 2. Register the wordtoolkit:// protocol link (points at the handler) ---
cmd = """" & Wsh.ExpandEnvironmentStrings("%WINDIR%") & "\\System32\\wscript.exe"" """ & _
      appDataDir & "\\sync-handler.vbs"" ""%1"""
On Error Resume Next
Wsh.RegWrite "HKCU\\Software\\Classes\\wordtoolkit\\shell\\open\\command\\", cmd, "REG_SZ"
If Err.Number <> 0 Then
    MsgBox "Could not register the wordtoolkit:// link." & vbCrLf & "Error: " & Err.Description, vbCritical, "Word Toolkit Setup"
    WScript.Quit
End If
On Error GoTo 0

' --- 3. Enable Word import access automatically (Trust access + all macros) ---
EnableWordAccess Wsh

' --- 4. Import connector modules into Word (handler is NOT imported) ---
On Error Resume Next
' Close any running Word instances to release locks on Normal.dotm
Set w = GetObject(, "Word.Application")
If Not w Is Nothing Then
    w.Quit
    WScript.Sleep 500
End If
Err.Clear
Wsh.Run "taskkill /F /IM winword.exe /T", 0, True

' Create a fresh, hidden Word instance
Set w = CreateObject("Word.Application")
If Not w Is Nothing Then
    w.Visible = False
    w.DisplayAlerts = 0
End If
On Error GoTo 0

If w Is Nothing Then
    MsgBox "Could not start Word application background process.", vbCritical, "Word Toolkit Setup"
    WScript.Quit
End If

Set normDoc = Nothing
Set comps = Nothing

On Error Resume Next
Set normDoc = w.NormalTemplate.OpenAsDocument()
If Not normDoc Is Nothing Then Set comps = normDoc.VBProject.VBComponents
On Error GoTo 0

If comps Is Nothing Then
    docCreated = False
    On Error Resume Next
    Set doc = w.Documents.Add()
    docCreated = True
    Set vbProj = w.NormalTemplate.VBProject
    If Not vbProj Is Nothing Then Set comps = vbProj.VBComponents
    If comps Is Nothing And Not doc Is Nothing Then Set comps = doc.VBProject.VBComponents
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
If docCreated And Not doc Is Nothing Then doc.Close False
w.NormalTemplate.Save
w.Quit
Set w = Nothing
On Error GoTo 0

MsgBox "Word Toolkit setup complete!" & vbCrLf & vbCrLf & _
       "• " & n & " file(s) downloaded from " & url & vbCrLf & _
       "• Word import access enabled automatically (Trust access + enable all macros)" & vbCrLf & _
       "• Registered wordtoolkit:// protocol for 1-click browser sync" & vbCrLf & _
       "• Connector modules successfully updated in Word (" & importedCount & " imported)" & vbCrLf & vbCrLf & _
       "You can now click '🚀 Export to Word' on the web tool anytime!", _
       vbInformation, "Word Toolkit Setup"

Sub EnableWordAccess(wShell)
    Dim verList, vKey, rPath, rPaths
    verList = Array("16.0", "15.0", "14.0", "12.0", "11.0")
    rPaths = Array( _
        "HKCU\\Software\\Microsoft\\Office\\", _
        "HKCU\\Software\\Policies\\Microsoft\\Office\\", _
        "HKLM\\Software\\Microsoft\\Office\\", _
        "HKLM\\Software\\Policies\\Microsoft\\Office\\", _
        "HKLM\\Software\\WOW6432Node\\Microsoft\\Office\\", _
        "HKLM\\Software\\WOW6432Node\\Policies\\Microsoft\\Office\\" _
    )
    For Each vKey In verList
        For Each rPath In rPaths
            On Error Resume Next
            wShell.RegWrite rPath & vKey & "\\Word\\Security\\AccessVBOM", 1, "REG_DWORD"
            wShell.RegWrite rPath & vKey & "\\Word\\Security\\Level", 1, "REG_DWORD"
            wShell.RegWrite rPath & vKey & "\\Word\\Security\\VBAWarnings", 1, "REG_DWORD"
            wShell.RegWrite rPath & vKey & "\\Word\\Security\\DisableAllMacros", 0, "REG_DWORD"
            wShell.RegWrite rPath & vKey & "\\Word\\Security\\ExtensionHardening", 0, "REG_DWORD"
            wShell.RegWrite rPath & vKey & "\\Word\\Options\\AccessVBOM", 1, "REG_DWORD"
            wShell.RegWrite rPath & vKey & "\\Common\\Security\\AccessVBOM", 1, "REG_DWORD"
            On Error GoTo 0
        Next
    Next
End Sub`,
'Enable_Word_Import_Access.vbs': `Option Explicit
' ============================================================
'  WORD TOOLKIT - ENABLE ALL WORD IMPORT ACCESS (standalone VBS)
' ============================================================
Dim Wsh, versions, v, secKey, successCount, errCount
Set Wsh = CreateObject("WScript.Shell")
versions = Array("16.0", "15.0", "14.0", "12.0", "11.0")
successCount = 0: errCount = 0
For Each v In versions
    secKey = "HKCU\\Software\\Microsoft\\Office\\" & v & "\\Word\\Security\\"
    On Error Resume Next
    Wsh.RegWrite secKey & "AccessVBOM", 1, "REG_DWORD"
    Wsh.RegWrite secKey & "Level", 1, "REG_DWORD"
    Wsh.RegWrite secKey & "VBAWarnings", 1, "REG_DWORD"
    Wsh.RegWrite secKey & "DisableAllMacros", 0, "REG_DWORD"
    Wsh.RegWrite secKey & "ExtensionHardening", 0, "REG_DWORD"
    If Err.Number = 0 Then successCount = successCount + 1 Else errCount = errCount + 1
    On Error GoTo 0
Next
MsgBox "Microsoft Word Import Access Granted!" & vbCrLf & vbCrLf & _
       "• Trust access to VBA project object model: ENABLED" & vbCrLf & _
       "• Macro execution security level: ENABLE ALL" & vbCrLf & _
       "• VBA macro warnings: DISABLED (Full Trust)" & vbCrLf & vbCrLf & _
       "Configured for " & successCount & " Microsoft Office version registry keys." & vbCrLf & _
       "If Word is currently open, please restart Word to apply the new permissions.", _
       vbInformation, "Word Toolkit - Access Manager"`,
'Reset_Word_Normal_Template.vbs': `Option Explicit
' ============================================================
'  WORD TOOLKIT - RESET & UNLOCK NORMAL TEMPLATE
' ============================================================
Dim Wsh, fso, templatesDir, normalFile, backupFile, w, versions, v, secKey, polKey, root, roots

Set Wsh = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

On Error Resume Next
Set w = GetObject(, "Word.Application")
If Not w Is Nothing Then
    w.Quit
    WScript.Sleep 500
End If
Err.Clear
Wsh.Run "taskkill /F /IM winword.exe /T", 0, True
On Error GoTo 0

templatesDir = fso.BuildPath(Wsh.ExpandEnvironmentStrings("%APPDATA%"), "Microsoft\\Templates")
normalFile = fso.BuildPath(templatesDir, "Normal.dotm")
backupFile = fso.BuildPath(templatesDir, "Normal_Backup_" & Replace(Replace(Replace(Now(), "/", "-"), ":", "-"), " ", "_") & ".dotm")

If fso.FileExists(normalFile) Then
    On Error Resume Next
    fso.CopyFile normalFile, backupFile, True
    fso.DeleteFile normalFile, True
    On Error GoTo 0
End If

versions = Array("16.0", "15.0", "14.0", "12.0", "11.0")
roots = Array( _
    "HKCU\\Software\\Microsoft\\Office\\", _
    "HKCU\\Software\\Policies\\Microsoft\\Office\\", _
    "HKLM\\Software\\Microsoft\\Office\\", _
    "HKLM\\Software\\Policies\\Microsoft\\Office\\", _
    "HKLM\\Software\\WOW6432Node\\Microsoft\\Office\\", _
    "HKLM\\Software\\WOW6432Node\\Policies\\Microsoft\\Office\\" _
)
For Each v In versions
    For Each root In roots
        On Error Resume Next
        Wsh.RegWrite root & v & "\\Word\\Security\\AccessVBOM", 1, "REG_DWORD"
        Wsh.RegWrite root & v & "\\Word\\Security\\Level", 1, "REG_DWORD"
        Wsh.RegWrite root & v & "\\Word\\Security\VBAWarnings", 1, "REG_DWORD"
        Wsh.RegWrite root & v & "\\Word\\Security\\DisableAllMacros", 0, "REG_DWORD"
        Wsh.RegWrite root & v & "\\Word\\Security\\ExtensionHardening", 0, "REG_DWORD"
        Wsh.RegWrite root & v & "\\Word\\Options\\AccessVBOM", 1, "REG_DWORD"
        Wsh.RegWrite root & v & "\\Common\\Security\\AccessVBOM", 1, "REG_DWORD"
        On Error GoTo 0
    Next
Next

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
       vbInformation, "Reset Word Template"`
};

/* ---------- DOM LOAD & EVENT LISTENERS ---------- */
document.addEventListener('DOMContentLoaded', () => {
  initTabs();
  initSettingsModal();
  initConnectorDownloads();
  initBatchDropZone();
  checkServerConnection();
});

function initTabs() {
  $$('#tabs button').forEach(btn => {
    btn.addEventListener('click', () => {
      $$('#tabs button').forEach(b => b.classList.remove('active'));
      btn.classList.add('active');
      $$('.tab-content').forEach(s => s.style.display = 'none');
      const target = $('#tab-' + btn.dataset.tab);
      if (target) target.style.display = 'block';
    });
  });
}

/* ---------- SERVER CONNECTIVITY & LOCAL FALLBACK ---------- */
async function checkServerConnection() {
  const dot = $('#status-dot');
  const text = $('#status-text');
  try {
    const res = await fetch(`${apiBaseUrl}/health`, { cache: 'no-store' });
    if (res.ok) {
      const data = await res.json();
      isServerOnline = true;
      dot.className = 'status-dot';
      text.textContent = 'Render Server Online';
    } else {
      throw new Error('Server returned error status');
    }
  } catch (e) {
    isServerOnline = false;
    dot.className = 'status-dot offline';
    text.textContent = 'Local Mode (Offline)';
  }
  refreshAllData();
}

function initSettingsModal() {
  const modal = $('#settings-modal');
  const input = $('#server-url-input');
  const autoSync = $('#set-auto-sync');

  $('#btn-open-settings').onclick = () => {
    input.value = apiBaseUrl;
    autoSync.checked = autoSyncEnabled();
    modal.style.display = 'flex';
  };
  $('#btn-close-settings').onclick = () => {
    modal.style.display = 'none';
  };
  $('#btn-save-settings').onclick = () => {
    let val = input.value.trim();
    if (!val) val = '/api';
    if (val.endsWith('/')) val = val.slice(0, -1);
    apiBaseUrl = val;
    localStorage.setItem('wt_server_url', apiBaseUrl);
    localStorage.setItem('wt_auto_sync', autoSync.checked ? '1' : '0');
    modal.style.display = 'none';
    showToast(autoSync.checked ? 'Server saved — auto-sync into Word is ON' : 'Server settings updated');
    checkServerConnection();
  };
}

// True by default; can be explicitly disabled in settings
function autoSyncEnabled() {
  const val = localStorage.getItem('wt_auto_sync');
  return val === null ? true : val === '1';
}

// If auto-sync is on, trigger a direct install for the given section group(s).
// Empty group = all groups of that section.
function maybeAutoSync(mac, sc, rb) {
  if (!autoSyncEnabled()) return;
  directInstallToWord(mac || '', sc || '', rb || '');
}

/* ---------- REFRESH DATA ---------- */
async function refreshAllData() {
  await Promise.all([refreshMacros(), refreshShortcuts(), refreshRibbon()]);
}

/* ---------- MACRO MANAGEMENT ---------- */
async function refreshMacros() {
  const listEl = $('#m-list');
  try {
    if (isServerOnline) {
      const res = await fetch(`${apiBaseUrl}/macros`);
      const data = await res.json();
      currentMacros = data.macros || [];
    } else {
      const raw = localStorage.getItem('wt_local_macros');
      currentMacros = raw ? JSON.parse(raw) : getSampleMacros();
    }
  } catch (e) {
    console.error('Failed to fetch macros:', e);
    currentMacros = getSampleMacros();
  }

  updateGroupSuggestionsAndFilter();
  renderGroupedMacros();
}

function getSampleMacros() {
  return [
    {
      id: 'm-sample-1',
      group: 'Formatting Tools',
      name: 'CleanFormatting',
      type: 'bas',
      code: `Attribute VB_Name = "CleanFormatting"\nOption Explicit\n\nSub RemoveExtraSpaces()\n    With Selection.Find\n        .ClearFormatting\n        .Replacement.ClearFormatting\n        .Text = "  "\n        .Replacement.Text = " "\n        .Forward = True\n        .Wrap = wdFindContinue\n        .Execute Replace:=wdReplaceAll\n    End With\n    MsgBox "Extra spaces removed!", vbInformation\nEnd Sub`,
      updatedAt: new Date().toISOString()
    }
  ];
}

function updateGroupSuggestionsAndFilter() {
  const datalist = $('#group-suggestions');
  const selectFilter = $('#m-filter-group');
  const selectedGroup = selectFilter.value;

  const groups = Array.from(new Set(currentMacros.map(m => m.group || 'General'))).sort();

  datalist.innerHTML = groups.map(g => `<option value="${escapeHtml(g)}">`).join('');

  selectFilter.innerHTML = '<option value="">All Groups</option>' + 
    groups.map(g => `<option value="${escapeHtml(g)}"${g === selectedGroup ? ' selected' : ''}>${escapeHtml(g)}</option>`).join('');
}

function renderGroupedMacros() {
  const listEl = $('#m-list');
  const searchQuery = ($('#m-search').value || '').toLowerCase().trim();
  const groupFilter = ($('#m-filter-group').value || '').toLowerCase().trim();

  let filtered = currentMacros.filter(m => {
    const groupName = (m.group || 'General').toLowerCase();
    const nameMatch = (m.name || '').toLowerCase().includes(searchQuery) ||
                      groupName.includes(searchQuery) ||
                      (m.code || '').toLowerCase().includes(searchQuery);
    const groupMatch = !groupFilter || groupName === groupFilter;
    return nameMatch && groupMatch;
  });

  if (filtered.length === 0) {
    listEl.innerHTML = `<div class="empty">No macros found. ${currentMacros.length === 0 ? 'Add a new macro above or import a .bas file.' : 'Try adjusting your search filter.'}</div>`;
    return;
  }

  // Group macros by Group Name
  const groupedMap = {};
  filtered.forEach(m => {
    const g = m.group || 'General';
    if (!groupedMap[g]) groupedMap[g] = [];
    groupedMap[g].push(m);
  });

  listEl.innerHTML = '';

  Object.keys(groupedMap).sort().forEach(groupName => {
    const items = groupedMap[groupName];
    const groupShortcuts = currentShortcuts.filter(sc =>
      (sc.group || 'General').toLowerCase() === groupName.toLowerCase()
    );
    const groupRibbons = currentRibbon.filter(rb =>
      (rb.group || 'General').toLowerCase() === groupName.toLowerCase()
    );

    const groupSec = document.createElement('div');
    groupSec.className = 'group-section';

    const header = document.createElement('div');
    header.className = 'group-header';
    header.innerHTML = `
      <div class="group-title-area">
        <h3 class="group-title">${escapeHtml(groupName)}</h3>
        <span class="group-count-badge">${items.length} macro${items.length === 1 ? '' : 's'}${groupShortcuts.length > 0 ? ` · ${groupShortcuts.length} shortcut set${groupShortcuts.length === 1 ? '' : 's'}` : ''}${groupRibbons.length > 0 ? ` · ${groupRibbons.length} ribbon` : ''}</span>
      </div>
      <button class="btn small secondary btn-export-group" data-group="${escapeHtml(groupName)}">Export Group (.bas)</button>
    `;

    const stampsDiv = document.createElement('div');
    stampsDiv.className = 'stamps';

    items.forEach(macro => {
      const stamp = document.createElement('div');
      stamp.className = 'stamp';
      stamp.innerHTML = `
        <div class="meta">
          <span class="badge">.${escapeHtml(macro.type || 'bas')}</span>
          <p class="name">${escapeHtml(macro.name)}</p>
          <span class="sub">Group: ${escapeHtml(macro.group || 'General')} · ${fmtDate(macro.updatedAt)}</span>
        </div>
        <div class="actions">
          <button class="btn small secondary" data-act="copy">Copy Code</button>
          <button class="btn small secondary" data-act="dl">Download</button>
          <button class="btn small secondary" data-act="edit">Edit</button>
          <button class="btn small ghost-danger" data-act="del">Delete</button>
        </div>
      `;

      stamp.querySelector('[data-act=copy]').onclick = () => {
        navigator.clipboard.writeText(macro.code);
        showToast(`Copied ${macro.name} to clipboard!`);
      };
      stamp.querySelector('[data-act=dl]').onclick = () => {
        downloadFile(`${macro.name}.${macro.type || 'bas'}`, macro.code);
      };
      stamp.querySelector('[data-act=edit]').onclick = () => {
        $('#m-group').value = macro.group || '';
        $('#m-shortcut-set').value = '';
        window.scrollTo({ top: $('#tab-macros').offsetTop, behavior: 'smooth' });
        showToast(`Loaded ${macro.name} — pick a shortcut set to recreate macros`);
      };
      stamp.querySelector('[data-act=del]').onclick = async () => {
        if (confirm(`Delete macro "${macro.name}" from group "${macro.group}"?`)) {
          await deleteMacro(macro.id);
        }
      };

      stampsDiv.appendChild(stamp);
    });

    // Bundled shortcut sets for this group
    groupShortcuts.forEach(sc => {
      const rows = (sc.csv.match(/\n/g) || []).length;
      const stamp = document.createElement('div');
      stamp.className = 'stamp shortcut-stamp';
      stamp.innerHTML = `
        <div class="meta">
          <span class="badge">.CSV</span>
          <p class="name">${escapeHtml(sc.name)}</p>
          <span class="sub">${rows} shortcut(s) · ${fmtDate(sc.updatedAt)}</span>
        </div>
        <div class="actions">
          <button class="btn small secondary" data-act="sdl">Download CSV</button>
          <button class="btn small ghost-danger" data-act="sdel">Delete</button>
        </div>
      `;
      stamp.querySelector('[data-act=sdl]').onclick = () => downloadFile(`${sc.name}.csv`, sc.csv);
      stamp.querySelector('[data-act=sdel]').onclick = async () => {
        if (confirm(`Delete shortcut set "${sc.name}" from group "${groupName}"?`)) {
          await deleteShortcut(sc.id);
        }
      };
      stampsDiv.appendChild(stamp);
    });

    // Bundled ribbon profiles for this group
    groupRibbons.forEach(rb => {
      const stamp = document.createElement('div');
      stamp.className = 'stamp ribbon-stamp';
      stamp.innerHTML = `
        <div class="meta">
          <span class="badge">.officeUI</span>
          <p class="name">${escapeHtml(rb.name)}</p>
          <span class="sub">${escapeHtml(rb.filename || 'Word.officeUI')} · ${fmtDate(rb.updatedAt)}</span>
        </div>
        <div class="actions">
          <button class="btn small secondary" data-act="rdl">Download</button>
          <button class="btn small ghost-danger" data-act="rdel">Delete</button>
        </div>
      `;
      stamp.querySelector('[data-act=rdl]').onclick = () => downloadBase64(rb.filename || 'Word.officeUI', rb.base64);
      stamp.querySelector('[data-act=rdel]').onclick = async () => {
        if (confirm(`Delete ribbon profile "${rb.name}" from group "${groupName}"?`)) {
          await deleteRibbon(rb.id);
        }
      };
      stampsDiv.appendChild(stamp);
    });

    header.querySelector('.btn-export-group').onclick = (e) => {
      e.stopPropagation();
      exportMacroGroup(groupName, items);
    };

    groupSec.appendChild(header);
    groupSec.appendChild(stampsDiv);
    listEl.appendChild(groupSec);
  });
}

// Search and filter listeners
$('#m-search').addEventListener('input', renderGroupedMacros);
$('#m-filter-group').addEventListener('change', renderGroupedMacros);

// Save Macro Click Event — creates macros from the chosen imported shortcut set
$('#m-save').addEventListener('click', async () => {
  const group = $('#m-group').value.trim() || 'General';
  const setName = $('#m-shortcut-set').value;
  const status = $('#m-status');

  if (!setName) {
    status.textContent = 'Select an imported shortcut set first.';
    status.className = 'status-msg err';
    return;
  }

  const set = currentShortcuts.find(s => s.name === setName);
  if (!set || !set.csv) {
    status.textContent = 'Shortcut set not found.';
    status.className = 'status-msg err';
    return;
  }

  // Extract unique commands from the CSV, one macro per command
  const commands = new Set();
  (set.csv.split(/\r?\n/).slice(1)).forEach(line => {
    if (!line.trim()) return;
    const parts = line.split(',');
    if (parts.length >= 2 && parts[1].trim()) commands.add(parts[1].trim());
  });

  if (commands.size === 0) {
    status.textContent = 'No commands found in this shortcut set.';
    status.className = 'status-msg err';
    return;
  }

  status.textContent = `Saving ${commands.size} macro(s)…`;
  status.className = 'status-msg';

  try {
    let saved = 0, replaced = 0;
    for (const cmd of commands) {
      const name = cmd.split('.').pop(); // Normal.Module3.DecrementLeading -> DecrementLeading
      const payload = {
        group,
        name,
        type: 'bas',
        code: `Attribute VB_Name = "${name}"\n\nSub ${name}()\n    ' Macro from shortcut set "${setName}"\nEnd Sub`
      };
      let action = 'created';
      if (isServerOnline) {
        const res = await fetch(`${apiBaseUrl}/macros`, {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify(payload)
        });
        const data = await res.json();
        if (!data.success) throw new Error(data.error || 'Server save failed');
        action = data.action;
      } else {
        const existingIdx = currentMacros.findIndex(m =>
          m.name.toLowerCase() === payload.name.toLowerCase() &&
          (m.group || 'General').toLowerCase() === group.toLowerCase()
        );
        const macroObj = {
          id: existingIdx >= 0 ? currentMacros[existingIdx].id : 'macro-local-' + Date.now(),
          ...payload, updatedAt: new Date().toISOString()
        };
        if (existingIdx >= 0) { currentMacros[existingIdx] = macroObj; action = 'replaced'; }
        else currentMacros.push(macroObj);
        localStorage.setItem('wt_local_macros', JSON.stringify(currentMacros));
      }
      if (action === 'replaced') replaced++; else saved++;
    }

    // Bundle the shortcut set into the same group (update in place, no duplicates)
    const scIdx = currentShortcuts.findIndex(s => s.name === setName);
    if (scIdx >= 0) {
      await saveShortcutItem({ id: currentShortcuts[scIdx].id, name: setName, group, csv: set.csv });
    } else {
      await saveShortcutItem({ name: setName, group, csv: set.csv });
    }

    showToast(`Saved ${saved} new macro(s), replaced ${replaced}`);
    status.textContent = `Saved ${saved + replaced} macro(s) into "${group}"!`;
    status.className = 'status-msg';
    refreshAllData();
    maybeAutoSync(group, '', '');
  } catch (e) {
    status.textContent = 'Save failed: ' + e.message;
    status.className = 'status-msg err';
  }
});

async function deleteMacro(id) {
  try {
    if (isServerOnline) {
      await fetch(`${apiBaseUrl}/macros/${id}`, { method: 'DELETE' });
    } else {
      currentMacros = currentMacros.filter(m => m.id !== id);
      localStorage.setItem('wt_local_macros', JSON.stringify(currentMacros));
    }
    showToast('Macro deleted');
    refreshMacros();
  } catch (e) {
    showToast('Could not delete macro: ' + e.message);
  }
}

function exportMacroGroup(groupName, items) {
  const bundledCode = items.map(m => `' ==========================================\n' Module: ${m.name}.${m.type}\n' Group: ${groupName}\n' ==========================================\n${m.code}`).join('\n\n');
  downloadFile(`${groupName.replace(/[^a-zA-Z0-9_]/g, '_')}_Macros.bas`, bundledCode);
  showToast(`Exported ${items.length} macro(s) from "${groupName}"`);
}

/* ---------- SHORTCUTS MANAGEMENT ---------- */
function populateMacroShortcutPicker() {
  const picker = $('#m-shortcut-set');
  if (!picker) return;
  const prev = picker.value;
  picker.innerHTML = '<option value="">— Select an imported shortcut set —</option>' +
    currentShortcuts.map(sc =>
      `<option value="${escapeHtml(sc.name)}"${sc.name === prev ? ' selected' : ''}>${escapeHtml(sc.name)}${sc.group && sc.group !== 'General' ? ` (${escapeHtml(sc.group)})` : ''}` 
    ).join('');
}

async function refreshShortcuts() {
  const listEl = $('#s-list');
  try {
    if (isServerOnline) {
      const res = await fetch(`${apiBaseUrl}/shortcuts`);
      const data = await res.json();
      currentShortcuts = data.shortcuts || [];
} else {
    const raw = localStorage.getItem('wt_local_shortcuts');
    let local = raw ? JSON.parse(raw) : [];
    const seen = new Map();
    local.forEach(s => {
      const key = (s.name || '').toLowerCase();
      const existing = seen.get(key);
      if (!existing || (s.updatedAt || '') > (existing.updatedAt || '')) seen.set(key, s);
    });
    currentShortcuts = Array.from(seen.values());
    localStorage.setItem('wt_local_shortcuts', JSON.stringify(currentShortcuts));
  }
  } catch (e) {
    currentShortcuts = [];
  }

  if (currentShortcuts.length === 0) {
    listEl.innerHTML = '<div class="empty">No shortcut sets saved yet.</div>';
    populateMacroShortcutPicker();
    return;
  }

  listEl.innerHTML = '';
  populateMacroShortcutPicker();
  currentShortcuts.forEach(sc => {
    const rows = (sc.csv.match(/\n/g) || []).length;
    const el = document.createElement('div');
    el.className = 'stamp';
    el.innerHTML = `
      <div class="meta">
        <span class="badge">.CSV</span>
        <p class="name">${escapeHtml(sc.name)}</p>
        <span class="sub">Group: ${escapeHtml(sc.group || 'General')} · ${rows} shortcut(s) · ${fmtDate(sc.updatedAt)}</span>
      </div>
      <div class="actions">
        <button class="btn small secondary" data-act="dl">Download</button>
        <button class="btn small ghost-danger" data-act="del">Delete</button>
      </div>`;
    el.querySelector('[data-act=dl]').onclick = () => downloadFile(`${sc.name}.csv`, sc.csv);
    el.querySelector('[data-act=del]').onclick = () => deleteShortcut(sc.id);
    listEl.appendChild(el);
  });
}

$('#s-save').addEventListener('click', async () => {
  const name = $('#s-name').value.trim();
  const group = $('#s-group').value.trim();
  const csv = $('#s-csv').value;
  const status = $('#s-status');

  if (!name || !csv.trim()) {
    status.textContent = 'Set name and CSV data are required.';
    status.className = 'status-msg err';
    return;
  }

  status.textContent = 'Saving shortcut set…';
  try {
    if (isServerOnline) {
      await fetch(`${apiBaseUrl}/shortcuts`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ name, group, csv })
      });
    } else {
      currentShortcuts.push({ id: 'sc-' + Date.now(), name, group: group || 'General', csv, updatedAt: new Date().toISOString() });
      localStorage.setItem('wt_local_shortcuts', JSON.stringify(currentShortcuts));
    }
    status.textContent = 'Saved!';
    status.className = 'status-msg';
    $('#s-name').value = ''; $('#s-csv').value = '';
    showToast('Saved shortcut set');
    refreshShortcuts();
    maybeAutoSync('', group, '');
  } catch (e) {
    status.textContent = 'Save failed: ' + e.message;
    status.className = 'status-msg err';
  }
});

async function deleteShortcut(id) {
  if (isServerOnline) {
    await fetch(`${apiBaseUrl}/shortcuts/${id}`, { method: 'DELETE' });
  } else {
    currentShortcuts = currentShortcuts.filter(s => s.id !== id);
    localStorage.setItem('wt_local_shortcuts', JSON.stringify(currentShortcuts));
  }
  showToast('Shortcut set deleted');
  refreshShortcuts();
}

/* ---------- RIBBON MANAGEMENT ---------- */
async function refreshRibbon() {
  const listEl = $('#r-list');
  try {
    if (isServerOnline) {
      const res = await fetch(`${apiBaseUrl}/ribbon`);
      const data = await res.json();
      currentRibbon = data.ribbon || [];
    } else {
      const raw = localStorage.getItem('wt_local_ribbon');
      currentRibbon = raw ? JSON.parse(raw) : [];
    }
  } catch (e) {
    currentRibbon = [];
  }

  if (currentRibbon.length === 0) {
    listEl.innerHTML = '<div class="empty">No ribbon customization profiles saved yet.</div>';
    return;
  }

  listEl.innerHTML = '';
  currentRibbon.forEach(rb => {
    const el = document.createElement('div');
    el.className = 'stamp';
    el.innerHTML = `
      <div class="meta">
        <span class="badge">.officeUI</span>
        <p class="name">${escapeHtml(rb.name)}</p>
        <span class="sub">Group: ${escapeHtml(rb.group || 'General')} · ${escapeHtml(rb.filename || 'Word.officeUI')} · ${fmtDate(rb.updatedAt)}</span>
      </div>
      <div class="actions">
        <button class="btn small secondary" data-act="dl">Download</button>
        <button class="btn small ghost-danger" data-act="del">Delete</button>
      </div>`;
    el.querySelector('[data-act=dl]').onclick = () => downloadBase64(rb.filename || 'Word.officeUI', rb.base64);
    el.querySelector('[data-act=del]').onclick = () => deleteRibbon(rb.id);
    listEl.appendChild(el);
  });
}

$('#r-save').addEventListener('click', () => {
  const name = $('#r-name').value.trim();
  const group = $('#r-group').value.trim();
  const file = $('#r-file').files[0];
  const status = $('#r-status');

  if (!name || !file) {
    status.textContent = 'Name and file are required.';
    status.className = 'status-msg err';
    return;
  }

  status.textContent = 'Reading & saving file…';
  const reader = new FileReader();
  reader.onload = async () => {
    try {
      const base64 = reader.result.split(',')[1];
      if (isServerOnline) {
        await fetch(`${apiBaseUrl}/ribbon`, {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ name, group, filename: file.name, base64 })
        });
      } else {
        await saveRibbonItem({ name, group: group || 'General', filename: file.name, base64 });
      }
      status.textContent = 'Saved!';
      status.className = 'status-msg';
      $('#r-name').value = ''; $('#r-group').value = ''; $('#r-file').value = '';
      showToast('Saved ribbon profile');
      refreshRibbon();
      maybeAutoSync('', '', group);
    } catch (e) {
      status.textContent = 'Save failed: ' + e.message;
      status.className = 'status-msg err';
    }
  };
  reader.readAsDataURL(file);
});

async function deleteRibbon(id) {
  if (isServerOnline) {
    await fetch(`${apiBaseUrl}/ribbon/${id}`, { method: 'DELETE' });
  } else {
    currentRibbon = currentRibbon.filter(r => r.id !== id);
    localStorage.setItem('wt_local_ribbon', JSON.stringify(currentRibbon));
  }
  showToast('Ribbon profile deleted');
  refreshRibbon();
}

/* ---------- IMPORT / EXPORT HUB ---------- */

/* ---- Progress bar helpers ---- */
function showProgress(label, percent) {
  const bar = $('#io-progress');
  if (!bar) return;
  bar.style.display = 'block';
  $('#io-progress-label').textContent = label;
  $('#io-progress-fill').style.width = Math.max(0, Math.min(100, percent)) + '%';
}

function hideProgress() {
  const bar = $('#io-progress');
  if (!bar) return;
  bar.style.display = 'none';
  $('#io-progress-fill').style.width = '0%';
}

// fetch wrapper that reports real byte-level upload/download progress
function fetchWithProgress(url, options = {}) {
  return new Promise((resolve, reject) => {
    const xhr = new XMLHttpRequest();
    xhr.open(options.method || 'GET', url, true);

    if (options.headers) {
      Object.keys(options.headers).forEach(k => xhr.setRequestHeader(k, options.headers[k]));
    }
    if (options.responseType) xhr.responseType = options.responseType;

    xhr.upload.onprogress = (e) => {
      if (options.onUpload && e.lengthComputable) {
        options.onUpload(Math.round((e.loaded / e.total) * 100));
      }
    };
    xhr.onprogress = (e) => {
      if (options.onDownload && e.lengthComputable) {
        options.onDownload(Math.round((e.loaded / e.total) * 100));
      }
    };
    xhr.onload = () => {
      if (xhr.status >= 200 && xhr.status < 300) resolve(xhr);
      else reject(new Error('Server returned status ' + xhr.status));
    };
    xhr.onerror = () => reject(new Error('Network error during transfer'));
    xhr.send(options.body || null);
  });
}

// Direct Word Sync bundle downloads
$('#io-sync-macros').addEventListener('click', async (e) => {
  e.preventDefault();
  if (isServerOnline) {
    try {
      showProgress('Downloading macro bundle…', 3);
      const xhr = await fetchWithProgress(`${apiBaseUrl}/sync/macros`, {
        onDownload: p => showProgress(`Downloading macro bundle… ${p}%`, p)
      });
      showProgress('Finalizing macro bundle…', 100);
      downloadFile('word-toolkit-macro-bundle.txt', xhr.responseText);
      showToast('Downloaded macro bundle');
    } catch (err) {
      showToast('Download failed: ' + err.message);
    } finally {
      hideProgress();
    }
  } else {
    const bundle = currentMacros.map(m => `@group=${m.group || 'General'}\n@name=${m.name}\n@type=${(m.type || 'bas').toLowerCase()}\n${m.code}\n#WTMACRO-END#\n`).join('');
    downloadFile('word-toolkit-macro-bundle.txt', 'WORDTOOLKIT MACRO BUNDLE v1\n' + bundle);
    showToast('Downloaded local macro bundle');
  }
});

$('#io-sync-shortcuts').addEventListener('click', async (e) => {
  e.preventDefault();
  if (isServerOnline) {
    try {
      showProgress('Downloading shortcut bundle…', 3);
      const xhr = await fetchWithProgress(`${apiBaseUrl}/sync/shortcuts`, {
        onDownload: p => showProgress(`Downloading shortcut bundle… ${p}%`, p)
      });
      showProgress('Finalizing shortcut bundle…', 100);
      downloadFile('word-toolkit-shortcut-bundle.csv', xhr.responseText);
      showToast('Downloaded shortcut bundle');
    } catch (err) {
      showToast('Download failed: ' + err.message);
    } finally {
      hideProgress();
    }
  } else {
    const sets = currentShortcuts.map(sc => `#SET:${sc.name}${sc.group && sc.group !== 'General' ? ` (${sc.group})` : ''}\n${sc.csv}\n`).join('');
    downloadFile('word-toolkit-shortcut-bundle.csv', sets);
    showToast('Downloaded local shortcut bundle');
  }
});

// Export to Word: generate a customized installer with the user's group selections
function populateExportWordGroups() {
  const macroGroups = [...new Set(currentMacros.map(m => (m.group || 'General').trim()).filter(Boolean))].sort();
  const shortcutGroups = [...new Set(currentShortcuts.map(s => (s.group || 'General').trim()).filter(Boolean))].sort();
  const ribbonGroups = [...new Set(currentRibbon.map(r => (r.group || 'General').trim()).filter(Boolean))].sort();
  const fill = (sel, groups) => {
    const s = $(sel);
    const prev = s.value;
    s.innerHTML = '<option value="">All groups</option>' +
      groups.map(g => `<option value="${escapeHtml(g)}"${g === prev ? ' selected' : ''}>${escapeHtml(g)}</option>`).join('');
    if (!groups.includes(prev)) s.value = '';
  };
  fill('#ew-macro-group', macroGroups);
  fill('#ew-shortcut-group', shortcutGroups);
  fill('#ew-ribbon-group', ribbonGroups);
}

$('#io-export-word').addEventListener('click', (e) => {
  e.preventDefault();
  populateExportWordGroups();
  $('#ew-server-url').value = apiBaseUrl;
  $('#export-word-modal').style.display = 'flex';
});

$('#btn-close-export-word').addEventListener('click', () => {
  $('#export-word-modal').style.display = 'none';
});

// One-time PC setup: imports connector into Word + registers wordtoolkit:// link
$('#btn-download-setup').addEventListener('click', () => {
  downloadFile('WordToolkit_Setup.vbs', CONNECTOR_FILES['WordToolkit_Setup.vbs']);
  showToast('Setup downloaded — close Word windows, then run it once on this PC');
});

// Direct install: trigger the registered wordtoolkit:// link, which tells the
// running Word to fetch the selection from the server and install it.
function directInstallToWord(mac, sc, rb) {
  const clean = v => (v || '').replace(/"/g, "'").replace(/[\r\n]+/g, ' ').trim();
  const url = clean($('#ew-server-url').value) || apiBaseUrl;
  const m = mac ? clean(mac) : '';
  const s = sc ? clean(sc) : '';
  const r = rb ? clean(rb) : '';
  const link = 'wordtoolkit://sync?u=' + encodeURIComponent(url) +
               '&m=' + encodeURIComponent(m) +
               '&s=' + encodeURIComponent(s) +
               '&r=' + encodeURIComponent(r);
  window.location.href = link;
  $('#export-word-modal').style.display = 'none';
  showToast('🎉 Exported to Word! Check Microsoft Word for the "Successfully Imported" message.');
}

$('#btn-export-direct-word').addEventListener('click', () => {
  directInstallToWord(
    $('#ew-macros').checked ? $('#ew-macro-group').value : '',
    $('#ew-shortcuts').checked ? $('#ew-shortcut-group').value : '',
    $('#ew-ribbon').checked ? $('#ew-ribbon-group').value : ''
  );
});

// One-click installs straight from each tab's saved list
$('#btn-install-macros').addEventListener('click', () => {
  directInstallToWord($('#m-filter-group').value, '', '');
});
$('#btn-install-shortcuts').addEventListener('click', () => {
  directInstallToWord('', $('#s-group').value, '');
});
$('#btn-install-ribbon').addEventListener('click', () => {
  directInstallToWord('', '', $('#r-group').value);
});

// One click: enable Word's import permissions via the registered link
$('#btn-enable-word-access').addEventListener('click', (e) => {
  e.preventDefault();
  window.location.href = 'wordtoolkit://sync?act=enable&u=' + encodeURIComponent(apiBaseUrl);
  showToast('Enabling Word import access… (if nothing happens, run the one-time setup once on this PC)');
});

$('#btn-generate-installer').addEventListener('click', () => {
  let src = CONNECTOR_FILES['Toolkit_Sync.bas'];
  const macroOn = $('#ew-macros').checked;
  const scOn = $('#ew-shortcuts').checked;
  const rbOn = $('#ew-ribbon').checked;
  if (!macroOn) src = src.replace('Private Const SYNC_MACROS As Boolean = True', 'Private Const SYNC_MACROS As Boolean = False');
  if (!scOn) src = src.replace('Private Const SYNC_SHORTCUTS As Boolean = True', 'Private Const SYNC_SHORTCUTS As Boolean = False');
  if (!rbOn) src = src.replace('Private Const SYNC_RIBBON As Boolean = True', 'Private Const SYNC_RIBBON As Boolean = False');
  const clean = v => (v || '').replace(/"/g, "'").replace(/[\r\n]+/g, ' ').trim();
  src = src
    .replace('@@SERVER_URL@@', clean($('#ew-server-url').value) || '/api')
    .replace('@@MACRO_GROUP@@', macroOn ? clean($('#ew-macro-group').value) : '')
    .replace('@@SHORTCUT_GROUP@@', scOn ? clean($('#ew-shortcut-group').value) : '')
    .replace('@@RIBBON_GROUP@@', rbOn ? clean($('#ew-ribbon-group').value) : '');
  downloadFile('WordToolkit_Install.bas', src);
  showToast('Installer generated — import it in Word and run SyncAllFromServer');
  $('#export-word-modal').style.display = 'none';
});

// Export full backup JSON
$('#io-export-all').addEventListener('click', async () => {
  if (isServerOnline) {
    try {
      showProgress('Downloading system backup…', 3);
      const xhr = await fetchWithProgress(`${apiBaseUrl}/export`, {
        onDownload: p => showProgress(`Downloading system backup… ${p}%`, p)
      });
      showProgress('Finalizing backup file…', 100);

      const disposition = String(xhr.getResponseHeader('Content-Disposition') || '');
      const match = disposition.match(/filename="?([^";]+)"?/);
      const filename = match ? match[1] : `word-toolkit-backup-${Date.now()}.json`;
      downloadFile(filename, xhr.responseText);
      showToast('Exported system backup package');
    } catch (err) {
      showToast('Export failed: ' + err.message);
    } finally {
      hideProgress();
    }
  } else {
    const backupPkg = {
      appName: 'Word Shortcuts & Macro Manager',
      version: '1.0',
      exportedAt: new Date().toISOString(),
      macros: currentMacros,
      shortcuts: currentShortcuts,
      ribbon: currentRibbon
    };
    downloadFile(`word-toolkit-backup-${Date.now()}.json`, JSON.stringify(backupPkg, null, 2));
    showToast('Exported system backup package');
  }
});

// Restore System Backup JSON
$('#io-btn-restore').addEventListener('click', () => {
  const file = $('#io-restore-file').files[0];
  const mode = $('#io-restore-mode').value;

  if (!file) {
    alert('Please select a system backup .json file first.');
    return;
  }

  const reader = new FileReader();
  reader.onload = async () => {
    try {
      const parsedData = JSON.parse(reader.result);
      if (isServerOnline) {
        showProgress('Uploading backup to server…', 3);
        const payload = JSON.stringify({ mode, data: parsedData });
        await fetchWithProgress(`${apiBaseUrl}/import`, {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: payload,
          onUpload: p => showProgress(`Uploading backup… ${p}%`, p)
        });
        showProgress('Import complete', 100);
        showToast('Backup restored on server');
      } else {
        // Offline Restore
        if (mode === 'replace') {
          currentMacros = parsedData.macros || [];
          currentShortcuts = parsedData.shortcuts || [];
          currentRibbon = parsedData.ribbon || [];
        } else {
          // Merge
          (parsedData.macros || []).forEach(newM => {
            const idx = currentMacros.findIndex(m => m.name.toLowerCase() === (newM.name || '').toLowerCase());
            if (idx >= 0) currentMacros[idx] = newM;
            else currentMacros.push(newM);
          });
        }
        localStorage.setItem('wt_local_macros', JSON.stringify(currentMacros));
        showToast('Local system backup restored!');
      }
      $('#io-restore-file').value = '';
      refreshAllData();
      maybeAutoSync('', '', '');
      hideProgress();
    } catch (e) {
      hideProgress();
      alert('Error parsing backup file: ' + e.message);
    }
  };
  reader.readAsText(file);
});

/* ---------- BATCH DRAG & DROP ZONE ---------- */

// Convert a Word-exported shortcut .txt (pipe format:
// KeyString|KeyCode|KeyCategory|Command|CommandParameter) into the tool CSV format.
function convertShortcutTxtToCsv(text) {
  const lines = (text || '').split(/\r?\n/).map(l => l.trim()).filter(Boolean);
  const headerIdx = lines.findIndex(l => /^(KeyString\||FORMAT:\s*KeyString\|)/i.test(l));
  const startIdx = headerIdx === -1 ? 0 : headerIdx + 1;

  const out = ['KeyCategory,Command,KeyCode,KeyCode2,KeyString'];
  let converted = 0;
  for (let i = startIdx; i < lines.length; i++) {
    const ln = lines[i];
    if (/^-{3,}/.test(ln) || /^Total:|^FORMAT:/i.test(ln) || /^KeyString\|/i.test(ln)) continue;
    const parts = ln.split('|');
    if (parts.length >= 4) {
      const keyString = parts[0].trim();
      const keyCode = parts[1].trim();
      const keyCategory = parts[2].trim();
      const command = parts[3].trim();
      if (keyString && /^-?\d+$/.test(keyCode) && /^-?\d+$/.test(keyCategory) && command) {
        out.push(`${keyCategory},${command},${keyCode},0,"${keyString.replace(/\"/g, '""')}"`);
        converted++;
      }
    }
  }
  if (converted === 0) return null;
  return { csv: out.join('\n'), converted };
}

// Read all dropped files, build a preview list, and show the modal.
async function handleBatchFiles(files) {
  const previews = [];
  for (const file of Array.from(files)) {
    const ext = file.name.split('.').pop().toLowerCase();
    const item = { file, name: file.name, ext };

    if (['bas', 'cls', 'frm'].includes(ext)) {
      const t = await file.text();
      let modName = file.name.replace(/\.[^/.]+$/, "");
      const m = t.match(/Attribute\s+VB_Name\s*=\s*"([^"]+)"/i);
      if (m && m[1]) modName = m[1];
      item.kind = 'macro';
      item.displayName = modName;
      item.sub = `${ext.toUpperCase()} module`;
      item.payload = { group: 'Imported Files', name: modName, type: ext, code: t };
    } else if (ext === 'csv') {
      const t = await file.text();
      item.kind = 'shortcut';
      item.displayName = file.name.replace(/\.csv$/i, '');
      item.sub = 'CSV shortcut set';
      item.payload = { name: file.name.replace(/\.csv$/i, ''), csv: t };
    } else if (ext === 'txt') {
      const t = await file.text();
      const conv = convertShortcutTxtToCsv(t);
      if (conv) {
        item.kind = 'shortcut';
        item.displayName = file.name.replace(/\.txt$/i, '');
        item.sub = `TXT → CSV (${conv.converted} shortcuts converted)`;
        item.payload = { name: file.name.replace(/\.txt$/i, ''), csv: conv.csv };
      } else {
        item.kind = 'skip';
        item.displayName = file.name;
        item.sub = 'TXT not in shortcut format — import as macro code instead?';
        item.payload = { name: file.name.replace(/\.txt$/i, ''), csv: t };
      }
    } else if (ext === 'officeui') {
      item.kind = 'ribbon';
      item.displayName = file.name.replace(/\.officeui$/i, '');
      item.sub = 'Ribbon / QAT profile';
    } else {
      item.kind = 'skip';
      item.displayName = file.name;
      item.sub = 'Unsupported file type';
    }
    previews.push(item);
  }
  showImportPreview(previews);
}

function showImportPreview(previews) {
  const list = $('#import-preview-list');
  list.innerHTML = '';
  previews.forEach((p, i) => {
    const icon = p.kind === 'macro' ? '🧩' : p.kind === 'shortcut' ? '⌨️' : p.kind === 'ribbon' ? '📎' : '⚠️';
    const sub = p.sub || '';
    const el = document.createElement('div');
    el.className = 'io-preview-item';
    el.innerHTML = `<span class="p-icon">${icon}</span><span class="p-name">${escapeHtml(p.displayName)}</span><span class="p-sub">${escapeHtml(sub)}</span>`;
    list.appendChild(el);
  });

  $('#import-preview-modal').style.display = 'flex';
  $('#btn-preview-cancel').onclick = () => {
    $('#import-preview-modal').style.display = 'none';
  };
  $('#btn-preview-start').onclick = async () => {
    $('#import-preview-modal').style.display = 'none';
    await runBatchImport(previews);
  };
}

async function runBatchImport(previews) {
  const total = previews.filter(p => p.kind !== 'skip').length;
  let count = 0;
  let failed = 0;
  const importedShortcutSets = [];
  showProgress(`Importing 0 of ${total} files…`, 0);

  for (let i = 0; i < previews.length; i++) {
    const p = previews[i];
    showProgress(`Importing "${p.displayName}"…`, Math.round((count / total) * 100));
    try {
      if (p.kind === 'macro') {
        await saveMacroItem(p.payload);
        count++;
      } else if (p.kind === 'shortcut') {
        await saveShortcutItem(p.payload);
        importedShortcutSets.push(p.payload);
        count++;
      } else if (p.kind === 'ribbon') {
        const base64 = await readFileAsBase64(p.file);
        await saveRibbonItem({ name: p.displayName, filename: p.file.name, base64 });
        count++;
      } else {
        failed++;
      }
    } catch (err) {
      failed++;
      console.error('Import failed for', p.displayName, err);
    }
    showProgress(`Imported ${count} of ${total} files…`, Math.round((count / total) * 100));
  }

  if (count === total) showProgress(`All ${count} file(s) imported`, 100);
  else showProgress(`Imported ${count} of ${total} (${failed} failed)`, Math.round((count / total) * 100));
  setTimeout(hideProgress, 1200);
  showToast(`Successfully processed ${count} file(s)!` + (failed > 0 ? ` ${failed} failed.` : ''));
  refreshAllData();
  maybeAutoSync('', '', '');

  // After import: ask whether to save the shortcut sets into the macro group
  const hasMacros = previews.some(p => p.kind === 'macro');
  if (importedShortcutSets.length > 0 && hasMacros) {
    showPostImportPrompt(importedShortcutSets);
  }
}

function showPostImportPrompt(shortcutSets) {
  $('#post-import-msg').textContent =
    `The import included ${shortcutSets.length} shortcut set(s). Save them into the same group as the imported macros?`;
  $('#post-import-modal').style.display = 'flex';

  $('#btn-post-skip').onclick = () => {
    $('#post-import-modal').style.display = 'none';
  };
  $('#btn-post-save').onclick = async () => {
    $('#post-import-modal').style.display = 'none';
    const group = $('#post-import-group').value.trim() || 'Imported Files';
    showProgress(`Saving ${shortcutSets.length} shortcut set(s) into "${group}"…`, 10);
    let done = 0;
    try {
      for (const sc of shortcutSets) {
        await saveShortcutItem({ name: sc.name, group, csv: sc.csv });
        done++;
        showProgress(`Saving ${sc.name}…`, Math.round((done / shortcutSets.length) * 100));
      }
      showToast(`Saved ${shortcutSets.length} shortcut set(s) into "${group}"`);
    } catch (err) {
      showToast('Could not save shortcut sets: ' + err.message);
    }
    setTimeout(hideProgress, 1200);
    refreshAllData();
    maybeAutoSync('', group, '');
  };
}

function initBatchDropZone() {
  const zone = $('#batch-drop-zone');
  const fileInput = $('#batch-file-input');

  zone.onclick = () => fileInput.click();

  zone.ondragover = (e) => { e.preventDefault(); zone.classList.add('dragover'); };
  zone.ondragleave = () => zone.classList.remove('dragover');
  zone.ondrop = (e) => {
    e.preventDefault();
    zone.classList.remove('dragover');
    if (e.dataTransfer.files.length) handleBatchFiles(e.dataTransfer.files);
  };

  fileInput.onchange = () => {
    if (fileInput.files.length) handleBatchFiles(fileInput.files);
  };
}

async function saveMacroItem(item) {
  if (isServerOnline) {
    await fetch(`${apiBaseUrl}/macros`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(item)
    });
  } else {
    const idx = currentMacros.findIndex(m => m.name.toLowerCase() === item.name.toLowerCase());
    const macroObj = { ...item, id: 'macro-local-' + Date.now(), updatedAt: new Date().toISOString() };
    if (idx >= 0) currentMacros[idx] = macroObj;
    else currentMacros.push(macroObj);
    localStorage.setItem('wt_local_macros', JSON.stringify(currentMacros));
  }
}

async function saveShortcutItem(item) {
  if (isServerOnline) {
    await fetch(`${apiBaseUrl}/shortcuts`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(item)
    });
  } else {
    const idx = currentShortcuts.findIndex(s =>
      (item.id && s.id === item.id) ||
      s.name.toLowerCase() === (item.name || '').toLowerCase()
    );
    if (idx >= 0) currentShortcuts[idx] = { ...currentShortcuts[idx], ...item, updatedAt: new Date().toISOString() };
    else currentShortcuts.push({ ...item, id: 'sc-' + Date.now(), updatedAt: new Date().toISOString() });
    localStorage.setItem('wt_local_shortcuts', JSON.stringify(currentShortcuts));
  }
}

// When a macro is saved to a group, also persist any shortcut set the user
// has prepared for that same group (bundled under the same group name).
async function saveGroupShortcutSet(group) {
  const scName = $('#s-name').value.trim();
  const scGroup = $('#s-group').value.trim();
  const scCsv = $('#s-csv').value;
  if (!scName || !scCsv.trim()) return;

  const targetGroup = scGroup || group || 'General';
  if ((scGroup || 'General').toLowerCase() !== (group || 'General').toLowerCase()) return;

  try {
    if (isServerOnline) {
      await fetch(`${apiBaseUrl}/shortcuts`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ name: scName, group: targetGroup, csv: scCsv })
      });
    } else {
      currentShortcuts.push({ id: 'sc-' + Date.now(), name: scName, group: targetGroup, csv: scCsv, updatedAt: new Date().toISOString() });
      localStorage.setItem('wt_local_shortcuts', JSON.stringify(currentShortcuts));
    }
    $('#s-name').value = ''; $('#s-group').value = ''; $('#s-csv').value = '';
    showToast(`Bundled shortcut set "${scName}" into "${targetGroup}"`);
    refreshShortcuts();
  } catch (e) {
    showToast('Shortcut set bundled, but sync failed: ' + e.message);
  }
}

async function saveRibbonItem(item) {
  if (isServerOnline) {
    await fetch(`${apiBaseUrl}/ribbon`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(item)
    });
  } else {
    const idx = currentRibbon.findIndex(r =>
      (item.id && r.id === item.id) ||
      (r.name.toLowerCase() === (item.name || '').toLowerCase() &&
       (r.group || 'General').toLowerCase() === ((item.group || 'General') || 'General').toLowerCase())
    );
    if (idx >= 0) currentRibbon[idx] = { ...currentRibbon[idx], ...item, updatedAt: new Date().toISOString() };
    else currentRibbon.push({ ...item, id: 'rb-' + Date.now(), updatedAt: new Date().toISOString() });
    localStorage.setItem('wt_local_ribbon', JSON.stringify(currentRibbon));
  }
}

function getActiveServerUrl() {
  let url = apiBaseUrl || '/api';
  if (!url.startsWith('http://') && !url.startsWith('https://')) {
    url = window.location.origin + (url.startsWith('/') ? url : '/' + url);
  }
  return url;
}

/* ---------- CONNECTOR DOWNLOADS ---------- */
function initConnectorDownloads() {
  const grid = $('#connector-downloads');
  grid.innerHTML = '';
  Object.keys(CONNECTOR_FILES).forEach(fname => {
    const btn = document.createElement('button');
    btn.className = 'btn secondary';
    btn.textContent = '↓ Download ' + fname;
    btn.onclick = () => {
      let content = CONNECTOR_FILES[fname];
      if (fname === 'WordToolkit_Setup.vbs') {
        content = content.replace('"http://localhost:3000/api"', `"${getActiveServerUrl()}"`);
      }
      downloadFile(fname, content);
    };
    grid.appendChild(btn);
  });
}

/* ---------- HELPERS ---------- */
function downloadFile(filename, content) {
  let finalContent = content;
  if (filename === 'WordToolkit_Setup.vbs' && typeof finalContent === 'string') {
    finalContent = finalContent.replace('"http://localhost:3000/api"', `"${getActiveServerUrl()}"`);
  }
  const blob = new Blob([finalContent], { type: 'text/plain;charset=utf-8' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url; a.download = filename;
  document.body.appendChild(a); a.click(); a.remove();
  URL.revokeObjectURL(url);
}

function downloadBase64(filename, base64Str) {
  const bin = atob(base64Str);
  const bytes = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
  const blob = new Blob([bytes], { type: 'application/octet-stream' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url; a.download = filename;
  document.body.appendChild(a); a.click(); a.remove();
  URL.revokeObjectURL(url);
}

function readFileAsBase64(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result.split(',')[1]);
    reader.onerror = reject;
    reader.readAsDataURL(file);
  });
}

function escapeHtml(str) {
  if (!str) return '';
  return String(str).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
}

function fmtDate(isoStr) {
  if (!isoStr) return '';
  try {
    return new Date(isoStr).toLocaleString(undefined, { dateStyle: 'medium', timeStyle: 'short' });
  } catch (e) {
    return '';
  }
}

function showToast(msg) {
  const existing = $('.toast');
  if (existing) existing.remove();
  const toast = document.createElement('div');
  toast.className = 'toast';
  toast.textContent = msg;
  document.body.appendChild(toast);
  setTimeout(() => {
    toast.style.opacity = '0';
    setTimeout(() => toast.remove(), 300);
  }, 2800);
}
