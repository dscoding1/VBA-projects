Attribute VB_Name = "Concatenation"
Sub concatenate()

Dim myrange As Range
Dim counter As Integer
Dim ColCounter As Integer
Dim CellLimit As Integer
Dim str1 As String
Dim str2 As String
Dim StrConcat As Range

Application.Calculation = xlCalculationManual
Application.ScreenUpdating = False

CellLimit = 200

For counter = 0 To 10
ColCounter = 0
    
Set myrange = Range("T2:X2").Offset(counter, 0)

Debug.Print myrange.Row

str1 = Range("T" & myrange.Row).Offset(0, ColCounter)

Set StrConcat = Range("AB" & myrange.Row)
    
    Do Until str1 = ""


    If Len(StrConcat) < CellLimit Then
    str1 = Range("T" & myrange.Row).Offset(0, ColCounter)
                
            If StrConcat = "" Then
            StrConcat = str1
            Else
                If Range("T" & myrange.Row).Offset(0, ColCounter).Offset(0, 1) = "" Then
                GoTo NextLine
            Else
        
            StrConcat = StrConcat & ", " & str1
        
                End If
            End If

    Else
    
    Set StrConcat = Range("AC" & myrange.Row)
    str1 = Range("T" & myrange.Row).Offset(0, ColCounter)
                
            If StrConcat = "" Then
            StrConcat = str1
            Else
                If Range("T" & myrange.Row).Offset(0, ColCounter).Offset(0, 1) = "" Then
                GoTo NextLine
            Else
        
            StrConcat = StrConcat & ", " & str1
        
                End If
            End If
    
        Debug.Print str1

ColCounter = ColCounter + 1
Debug.Print Len(StrConcat)
    
        End If
        
    Loop

Debug.Print myrange.Address

NextLine:
    
Next counter

Application.Calculation = xlCalculationManual
Application.ScreenUpdating = True

End Sub
