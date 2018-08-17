Attribute VB_Name = "Module3"
Sub Rename()

Dim Source As Range
Dim OldFile As String
Dim NewFile As String

Set Source = Cells(1, 1).CurrentRegion

For Row = 1 To Source.Rows.Count
    OldFile = ActiveSheet.Cells(Row, 1)
    NewFile = ActiveSheet.Cells(Row, 2)

    ' rename files
    Name OldFile As NewFile

Next

End Sub


Sub MoveFiles()
    Dim FSO As Object
    Dim SourceFileName, DestinFileName As String
    
    For counter = 0 To 31
    On Error Resume Next
    Debug.Print counter
    SourceFileName = Cells(1, 1).Offset(counter, 0).Value
    DestinFileName = Cells(1, 2).Offset(counter, 0).Value
    
    Set FSO = CreateObject("Scripting.Filesystemobject")
    SourceFileName = "\\Hlrothgfc3ds\Global\Property records management\Tickets\SPLIT\" & Cells(1, 1).Offset(counter, 0).Value & ".xlsx"
    DestinFileName = "\\Hlrothgfc3ds\Global\Property records management\Tickets\" & Cells(1, 2).Offset(counter, 0).Value & "\"

    FSO.MoveFile Source:=SourceFileName, Destination:=DestinFileName

Next counter
    
    MsgBox (SourceFileName + " Moved to " + DestinFileName)
    

End Sub
