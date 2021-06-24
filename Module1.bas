Attribute VB_Name = "Module1"
Sub TickerInformation()

    Dim Ticker As String
    Dim UniqueIndex As Double
    Dim OpeningPrice As Double
    Dim ClosingPrice As Double
    Dim YearlyChange As Double
    Dim YearlyPercent As Double
    Dim TickerVolume As Double
    
    ' loop the calumn functions through all the sheets in the workbook
    For Each Worksheet In Worksheets
        Worksheet.Activate ' Mark work sheet as active for ActiveSheet methods.
        
        ' This is an alternative to cells(row,column), I used range because it was quicker to reference the letter of the columns
        ' Set generated columns headers for every sheet
        Range("I1") = "<ticker>"
        Range("J1") = "<yearly change>"
        Range("K1") = "<percent change>"
        Range("L1") = "<total volume>"
        
        ' Set generated row headers for bonus table
        Range("N2") = "Greastest % Increase"
        Range("N3") = "Greatest % Decrease"""
        Range("N4") = "Greatest Total Volume"
        
        Dim LastRowA As Double '
        
        ' Determine the index of the last used row in column "A"
        LastRowA = Worksheet.Cells(Rows.Count, "A").End(xlUp).Row
        
        ' Reset the variables for the new sheet
        
        Ticker = ""
        UniqueIndex = 2
        
        YearlyChange = 0 ' something
        YearlyPercent = 0
        TickerVolume = 0
        ClosingPrice = 0
        
        ' Set opening price for the first iteration
        OpeningPrice = Range("F2")
        
        For i = 2 To LastRowA
            
            ' Increase ticker volume for each iteration through the rows
            TickerVolume = TickerVolume + Range("G" & i).Value
            
            ' Make sure there is a positive value for opening price, if it is 0, continue to iterate until a value exists
            If OpeningPrice = 0 Then
            
                OpeningPrice = Range("C" & i) 'Change open price to next unique ID
                
            End If
            
            
            If Cells(i + 1, 1) <> Cells(i, 1) Then
                
                ' extract the unique name from Column A and put it into Column I
                ' create an indexer for Unique Ticker
                
                Ticker = Range("A" & i) ' pulling unique ticker
                Range("I" & UniqueIndex) = Range("A" & i) ' populate column I with unique ticker
                
                ClosingPrice = Range("F" & i) ' pull closing price
                
                YearlyChange = ClosingPrice - OpeningPrice ' Calculate yearly change
                Range("J" & UniqueIndex) = YearlyChange 'Populate yearly percent change
                
                ' Calculate yearly percent change
                ' To prevent opening being 0
                If OpeningPrice = 0 Then
                
                    YearlyPercent = 0
                
                Else
                    
                    YearlyPercent = (YearlyChange / OpeningPrice)
                
                End If
                
                
                
                Range("K" & UniqueIndex) = YearlyPercent 'Populate yearly percent change
                Range("K" & UniqueIndex).NumberFormat = "0.00%"
                
                ' Color code yearly change
                If YearlyPercent > 0 Then
                    Range("J" & UniqueIndex).Interior.ColorIndex = 4
                Else
                    Range("J" & UniqueIndex).Interior.ColorIndex = 3
                End If
                
                Range("L" & UniqueIndex) = TickerVolume 'Populate Ticker Volume
                
                'iterate UniqueIndex
                UniqueIndex = UniqueIndex + 1
                
                ' Initializing ticker volume to 0 so each unique ticker can reset
                ' Reset ticker volume
                TickerVolume = 0
                
                ' Update opening price for the next iteration.
                OpeningPrice = Range("C" & i + 1)
                
            End If
            
        Next i
        
        'list the return variable
        Dim GreatestValue As Double
        Dim GreatestPercent As Double
        Dim GreatestDecrease As Double
        Dim GreatestDecreaseTicker As String
        Dim GreatestPercentTicker As String
        Dim GreatestVolumeTicker As String
        
        'reset return variable for each worksheet to 0
        GreatestDecrease = 0
        GreatestPercent = 0
        GreatestVolume = 0
        UniqueValue = UniqueValue - 1 'Set the Index to the length of the newly generated columns
        
        For Index = 2 To UniqueIndex
        
            If GreatestDecrease > Range("K" & Index) Then
                GreatestDecrease = Range("K" & Index)
                GreatestDecreaseTicker = Range("I" & Index)
            ElseIf GreatestVolume < Range("L" & Index) Then
                GreatestVolume = Range("L" & Index)
                GreatestVolumeTicker = Range("I" & Index)
            ElseIf GreatestPercent < Range("K" & Index) Then
                GreatestPercent = Range("K" & Index)
                GreatestPercentTicker = Range("I" & Index)
            End If
        
        Next Index
        
        'Populate the cells
        Range("O2") = GreatestPercentTicker
        Range("O3") = GreatestDecreaseTicker
        Range("O4") = GreatestVolumeTicker
        
    Next Worksheet

End Sub
