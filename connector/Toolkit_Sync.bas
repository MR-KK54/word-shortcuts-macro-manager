Attribute VB_Name = "Toolkit_Sync"
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

    MsgBox "🎉 Successfully Imported into Microsoft Word!" & vbCrLf & vbCrLf & _
           "• " & macMsg & vbCrLf & _
           "• " & scMsg & vbCrLf & _
           "• " & rbMsg & vbCrLf & vbCrLf & _
           "All items are now imported and ready for immediate use in Word!", _
           vbInformation, "Microsoft Word - Import Successful"
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

    MsgBox "🎉 Successfully Imported into Microsoft Word!" & vbCrLf & vbCrLf & _
           "• " & macMsg & vbCrLf & _
           "• " & scMsg & vbCrLf & _
           "• " & rbMsg & vbCrLf & vbCrLf & _
           "All items are now imported and ready for immediate use in Word!", _
           vbInformation, "Microsoft Word - Import Successful"
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
    On Error Resume Next
    Set vbProj = Application.NormalTemplate.VBProject
    If vbProj Is Nothing Then Set vbProj = Application.VBE.ActiveVBProject
    On Error GoTo HandleErr

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
                tmpFile = Environ$("TEMP") & "\wtk_" & compName & "." & compType
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
    folderPath = Environ$("APPDATA") & "\Microsoft\Office"
    If Not fso.FolderExists(folderPath) Then fso.CreateFolder folderPath
    filePath = folderPath & "\Word.officeUI"

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
End Function