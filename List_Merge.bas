Attribute VB_Name = "Module1"
Sub ListMerge()

Dim shmerge As Worksheet
Dim sht As Worksheet
Dim LastRow As Long

Set shmerge = Worksheets("Merges")
Set sht = Worksheets("Merges Change")

Dim counter As Integer
Dim colCount As Integer

sht.Select

For colCount = 0 To 123

    For counter = 0 To shmerge.Range("L4").Offset(colCount, 0) - 1

LastRow = sht.Cells(sht.Rows.Count, "A").End(xlUp).Row

        shmerge.Range("C4").Offset(colCount, counter).Copy sht.Range("A" & LastRow + 1)
        shmerge.Range("B4").Offset(colCount, 0).Copy sht.Range("B" & LastRow + 1)
        
         
        Debug.Print LastRow
        
    Next counter
    

Next colCount

sht.Select

End Sub
