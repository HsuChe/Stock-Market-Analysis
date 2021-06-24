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
        
        ' Set generated columns headers for every sheet
        Range("I1") = "<ticker>"
        Range("J1") = "<yearly change>"
        Range("K1") = "<percent change>"
        Range("L1") = "<total volume>"
        Range("N2") = "Greastest % Increase"
        Range("N3") = "Greatest % Decrease"""
        Range("N4") = "Greatest Total Volume"
        
        Dim LastRowA As Double '
        LastRowA = Worksheet.Cells(Rows.Count, "A").End(xlUp).Row ' Determine the index of the last used row in column "A"
        
        'Reset the variables for the new sheet
        
        Ticker = ""
        UniqueIndex = 2
        OpeningPrice = 0
        ClosingPrice = Range("F2")
        YearlyChange = 0
        YearlyPercent = 0
        TickerVolume = 0
        
        For i = 2 To LastRowA
            
            TickerVolume = TickerVolume + Range("G" & i).Value
            If OpeningPrice = 0 Then
            
                OpeningPrice = Range("C" & i) 'Change open price to next unique ID
                
            End If
            
0
            If Cells(i + 1, 1) <> Cells(i, 1) Then
                
                ' extract the unique name from Column A and put it into Column I
                ' create an indexer for Unique Ticker
                
                Ticker = Range("A" & i) ' pulling unique ticker
                Range("I" & UniqueIndex) = Range("A" & i) ' populate column I with unique ticker
                
                ClosingPrice = Range("F" & i) ' pull closing price
                
                YearlyChange = ClosingPrice - OpeningPrice ' Calculate yearly change
                Range("J" & UniqueIndex) = YearlyChange 'Populate yearly percent change
                
                ' Calculate yearly percent change
                If OpeningPrice = 0 Then
                
                    YearlyPercent = 0
                
                Else
                    
                    YearlyPercent = (YearlyChange / OpeningPrice)
                
                End If
                
                
                Range("K" & UniqueIndex) = YearlyPercent 'Populate yearly percent change
                
                ' Color code yearly change
                If YearlyPercent > 0 Then
                    Range("J" & UniqueIndex).Interior.ColorIndex = 4
                Else
                    Range("J" & UniqueIndex).Interior.ColorIndex = 3
                End If
                
                Range("L" & UniqueIndex) = TickerVolume 'Populate Ticker Volume
                
                'iterate UniqueIndex
                UniqueIndex = UniqueIndex + 1
                
                'Reset ticker volume
                TickerVolume = 0
                'Reset Opening Price
                OpeningPrice = 0
                
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
