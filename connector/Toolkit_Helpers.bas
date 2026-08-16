Attribute VB_Name = "Toolkit_Helpers"
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
