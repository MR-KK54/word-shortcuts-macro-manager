Attribute VB_Name = "Toolkit_Menu"
Option Explicit

Sub ShowToolkitMenu()
    Dim choice As String, msg As String
    msg = "WORD CUSTOMIZATION TOOLKIT" & vbCrLf & vbCrLf & _
          "  1 - Export Macros (To Files)" & vbCrLf & _
          "  2 - Import Macros (Replace Existing)" & vbCrLf & _
          "  3 - Export Keyboard Shortcuts (.CSV)" & vbCrLf & _
          "  4 - Import Keyboard Shortcuts (.CSV)" & vbCrLf & _
          "  5 - Export Ribbon / QAT Customizations" & vbCrLf & _
          "  6 - Import Ribbon / QAT Customizations" & vbCrLf & vbCrLf & _
          "Enter a option number (1-6) or click Cancel:"
    choice = InputBox(msg, "Word Customization Toolkit")
    Select Case choice
        Case "1": ExportAllMacros
        Case "2": ImportAllMacros
        Case "3": ExportKeyboardShortcuts
        Case "4": ImportKeyboardShortcuts
        Case "5": ExportRibbonCustomizations
        Case "6": ImportRibbonCustomizations
        Case ""
        Case Else
            MsgBox "Please enter a number from 1 to 6.", vbExclamation, "Word Customization Toolkit"
    End Select
End Sub
