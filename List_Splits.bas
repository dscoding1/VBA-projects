Attribute VB_Name = "Module2"
Sub ListSplits()

Dim SplitList As Worksheet
Dim wsplits As Worksheet
Dim LastRow As Long

Set wsplits = Worksheets("NEW SPLITS")
Set SplitList = Worksheets("NEW SPLITS LIST")

Dim counter As Integer
Dim colCount As Integer

SplitList.Select

For colCount = 0 To 65

    For counter = 0 To wsplits.Range("P4").Offset(colCount, 0) - 1

LastRow = SplitList.Cells(SplitList.Rows.Count, "A").End(xlUp).Row

        wsplits.Range("C4").Offset(colCount, counter).Copy SplitList.Range("B" & LastRow + 1)
        wsplits.Range("B4").Offset(colCount, 0).Copy SplitList.Range("A" & LastRow + 1)
        
         
        Debug.Print LastRow
        
    Next counter
    

Next colCount

SplitList.Select

End Sub

