Attribute VB_Name = "Module1"
Dim rngStart As Excel.Range
Dim colCount As Long
Dim rngDestinationStart As Excel.Range

Sub test()
    Dim i As Long, rr As Long
    Dim counter As Integer
    
    
    RowCount = Range("A" & Rows.Count).End(xlUp).Row
    
    Cells(RowCount + 1, 1).Value = "End"
    Cells(RowCount + 1, 1).Font.Bold = True
 
For i = 5 To RowCount
    Set rngStart = Range("A" & i)
    
        If Range("A" & i).Value = "End" Then
        GoTo 1
        Else
    
    If rngStart.Font.Bold = False Then
    
         'Range("A" & i).Font.Bold = False
         'Range("C" & Range("C" & Rows.Count).End(xlUp).Offset(1, 1))
        
        colCount = NonBolds(rngStart.Row, rngStart.Column)
        Set rngDestinationStart = Range("F" & Rows.Count).End(xlUp).Offset(1, 0)
        'Range("A" & i).Copy Destination:=Range("C" & Range("C" & Rows.Count).End(xlUp).Offset(1, 1).Row)
        For j = 0 To colCount - 1
            'Range("A" & i + j).Copy Destination:=Range("C" & Range("C" & Rows.Count).End(xlUp).Offset(1, j).Row)
            Range("A" & i + j).Copy Destination:=rngDestinationStart.Offset(0, j)
        Next j
        i = i + colCount - 1
       
    Else
        
    If Range("A" & i).Font.Bold = True Then
            Range("A" & i).Copy Destination:=Range("E" & Range("E" & Rows.Count).End(xlUp).Offset(1, 0).Row)
            
        End If
        End If
        End If
        
    Next i
    
1

MsgBox ("It's Done")

End Sub

Function NonBolds(i As Long, j As Long) As Long
' i and j give the starting cell coordinates. j will always be the same
   ' i = i + 1
   Dim thiscounter As Long
   thiscounter = 0
    Do While Cells(i, j).Font.Bold = False And IsEmpty(ActiveCell) = False
    

        Debug.Print Cells(i, j).Address
        i = i + 1
        thiscounter = thiscounter + 1
        
        If j > 50 Then Exit Function
        
    Loop
    NonBolds = thiscounter
End Function

