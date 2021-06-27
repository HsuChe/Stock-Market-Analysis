Attribute VB_Name = "Module1"
Sub TickerInformation()

    Dim Ticker As String
    Dim UniqueIndex As Double
    Dim OpeningPrice As Double
    Dim ClosingPrice As Double
    Dim YearlyChange As Double
    Dim YearlyPercent As Double
    Dim TickerVolume As Double
    
    ' Create a for loop for each sheet/year of the workbook
    For Each Worksheet In Worksheets
        
        ' Mark work sheet as active to use the macro on the sheet.
        Worksheet.Activate
        
        ' Set generated columns headers for every sheet
        ' This is an alternative to cells(row,column), I used range method because it was quicker to reference the letter of the columns instead of counting the columns for cells
        Range("I1") = "<ticker>"
        Range("J1") = "<yearly change>"
        Range("K1") = "<percent change>"
        Range("L1") = "<total volume>"
        
        ' Set generated row headers for bonus table
        Range("N2") = "Greastest % Increase"
        Range("N3") = "Greatest % Decrease"
        Range("N4") = "Greatest Total Volume"
        Range("O1") = "Ticker"
        Range("P1") = "Value"
        
        ' Create a variable for the last used row of column A
        Dim LastRowA As Double
        
        ' Set value for the last used row for column A
        LastRowA = Worksheet.Cells(Rows.Count, "A").End(xlUp).Row
        
        ' Reset the variables for the new sheet
        Ticker = ""
        UniqueIndex = 2
        YearlyChange = 0
        YearlyPercent = 0
        TickerVolume = 0
        ClosingPrice = 0
        
        ' Set opening price for the first iteration because the for loop will be taking the last iteration of each unique variable.
        OpeningPrice = Range("F2")
        
        ' Start the for loop to iterate through each row of column A
        For i = 2 To LastRowA
            
            ' Increase ticker volume for each iteration through the rows, value will be returned to column L when conditional loop triggers
            TickerVolume = TickerVolume + Range("G" & i).Value
            
            ' We need to make sure that the opening price from the conditional loop is not 0, if it is 0, then we will set a new opening price the first time it is not 0
            If OpeningPrice = 0 Then
            
                ' Take the current value of opening price (column C) if the opening price in memory is 0
                OpeningPrice = Range("C" & i)
                
            End If
            
            ' Begin the conditional when a cell in column A has a different value to the next iteration, this will allow us to find the index where the switch happens
            ' I am comparing the current index to the next one through +1, so we will be getting the last iteration of an unique ticker name
            If Cells(i + 1, 1) <> Cells(i, 1) Then
                
                ' Extract ticker name from column A to be returned to Column I
                Ticker = Range("A" & i)
                ' Populate Column I with ticker.
                Range("I" & UniqueIndex) = Range("A" & i) ' populate column I with unique ticker
                
                ' Pull the closing price from the current index i to be used in calculation for change year on year
                ClosingPrice = Range("F" & i) ' pull closing price
                
                ' Calculate the yearly change for the ticker.
                YearlyChange = ClosingPrice - OpeningPrice
                ' Return YearlyChange to Column J
                Range("J" & UniqueIndex) = YearlyChange
                
                ' Calculate yearly percent change
                ' Test to see if Opening Price is 0 if the entire ticker has 0 for opening price, which will slip pass the first test
                If OpeningPrice = 0 Then
                    
                    ' If the entire opening price for the ticker is 0, we set the percent change to be 0
                    YearlyPercent = 0
                
                Else
                    
                    ' If the opening price is not 0, we calculate percent change
                    YearlyPercent = (YearlyChange / OpeningPrice)
                
                End If
                
                
                ' Return YearlyPercent to column K
                Range("K" & UniqueIndex) = YearlyPercent
                ' Format column k to percent with 2 decimal places
                Range("K" & UniqueIndex).NumberFormat = "0.00%"
                
                ' Color code yearly change
                If YearlyPercent > 0 Then
                    Range("J" & UniqueIndex).Interior.ColorIndex = 4
                Else
                    Range("J" & UniqueIndex).Interior.ColorIndex = 3
                End If
                
                'Return total volume to column L
                Range("L" & UniqueIndex) = TickerVolume
                
                ' Iterate UniqueIndex to track the current index in the summary table.
                UniqueIndex = UniqueIndex + 1
                
                ' Initializing ticker volume to 0 so each unique ticker can reset
                TickerVolume = 0
                
                ' Update opening price for the next iteration. The value will be fed back into the for loop.
                OpeningPrice = Range("C" & i + 1)
                
            End If
            
        Next i
        
        'list the return variable for the bonus question
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
        
        'Set the Index to the length of the newly generated columns from previous for loop
        UniqueValue = UniqueValue - 1
        
        ' Begin for loop to iterate through each row of the newly generated columns
        For Index = 2 To UniqueIndex
            
            ' Compare GreatestDecrease(starts at 0) to the cell value at current index, this will take the smallest value from the column
            If GreatestDecrease > Range("K" & Index) Then
                GreatestDecrease = Range("K" & Index)
                GreatestDecreaseTicker = Range("I" & Index)
            ' Compare GreatestIncrease(starts at 0) to the cell value at current index, this will take the largest value from the column
            ElseIf GreatestVolume < Range("L" & Index) Then
                GreatestVolume = Range("L" & Index)
                GreatestVolumeTicker = Range("I" & Index)
            ' Compare Greatest percent change(starts at 0) to the cell value at current index, this will take the largest value from the column
            ElseIf GreatestPercent < Range("K" & Index) Then
                GreatestPercent = Range("K" & Index)
                GreatestPercentTicker = Range("I" & Index)
            End If
        
        Next Index
        
        ' Return all the ticker names for the bonus columns.
        Range("O2") = GreatestPercentTicker
        Range("O3") = GreatestDecreaseTicker
        Range("O4") = GreatestVolumeTicker
        ' Return all the values for the bonus columns.
        Range("P2") = GreatestPercent
        Range("P3") = GreatestDecrease
        Range("P4") = GreatestVolume
        
    Next Worksheet

End Sub
