Attribute VB_Name = "Module1"
Sub GetSheets()

 Dim fldr As FileDialog
    Dim sItem As String
    Set fldr = Application.FileDialog(msoFileDialogFolderPicker)
    With fldr
        .Title = "Select a Folder"
        .AllowMultiSelect = False
        .InitialFileName = Application.DefaultFilePath
        If .Show <> -1 Then GoTo NextCode
        sItem = .SelectedItems(1)
    End With
NextCode:
    GetFolder = sItem
    Set fldr = Nothing

MsgBox (sItem)

Path = sItem & "\"

Filename = Dir(Path & "*.xls")
  Do While Filename <> ""
  Workbooks.Open Filename:=Path & Filename, ReadOnly:=True
     For Each Sheet In ActiveWorkbook.Sheets
     Sheet.Copy After:=ThisWorkbook.Sheets(1)
  Next Sheet
     Workbooks(Filename).Close
     Filename = Dir()
  Loop
End Sub
