Attribute VB_Name = "Module1"
Sub CreateSheetsFromAList()
'UpdatebyKutoolsforExcel20150916
    Dim Rg As Range
    Dim Rg1 As Range
    Dim xAddress As String
    On Error Resume Next
    xAddress = Application.ActiveWindow.RangeSelection.Address
    Set Rg = Application.InputBox("Select a range:", "Kutools for Excel", , , , , , 8)
    If Rg Is Nothing Then Exit Sub
    For Each Rg1 In Rg
        If Rg1 <> "" Then
            Call Sheets.Add(, Sheets(Sheets.Count))
            Sheets(Sheets.Count).Name = Rg1.Value
        End If
    Next
End Sub
