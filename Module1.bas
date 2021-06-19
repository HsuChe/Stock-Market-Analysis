Attribute VB_Name = "Module1"
Sub UniqueTicker()
    
    
    Columns("A:A").Select
    Selection.Copy
    Columns("I:I").Select
    ActiveSheet.Paste
    ActiveSheet.Range("I:I").RemoveDuplicates Columns:=1, Header:=xlNo


End Sub


Sub TickerInformation()

    ' we need to get the max index for used cells from column A
    Dim LastRowA As Long
    LastRowA = ActiveSheet.Cells(Rows.Count, "A").End(xlUp).Row
    
    Dim FirstIndex() As Long
    Dim LastIndex() As Long
    Dim IndexCount As Long
    Dim UniqueIndex As Long
    Dim OpeningPrice As Double
    Dim ClosingPrice As Double
    Dim YearlyChange As Double
    
    UniqueIndex = 0
    
    For i = 1 To LastRowA
    
        
        If Cells(i, 1) <> Cells(i + 1, 1) Then
            
            ' extract the unique name from Column A and put it into Column I
            ' create an indexer for Unique Ticker
            UniqueIndex = UniqueIndex + 1
            Range("I" & UniqueIndex) = Range("A" & i) 'pulling the last ticker from their set
            Range("J" & UniqueIndex) = i - 261 ' OpeningPrice
            Range("K" & UniqueIndex) = i ' ClosingPrice
            
            
        End If
        
    Next i


End Sub
