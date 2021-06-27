
<br />
<p align="center">

  <h3 align="center">Stock Ticker Summary Table</h3>

  <p align="center">
     An explanation for the creation of a summary table from ticker information
    <br />
    <a href="https://github.com/HsuChe/VBA_challenge"><strong>Project Github URL »</strong></a>
    <br />
    <br />
  </p>
</p>



<!-- ABOUT THE PROJECT -->
## About The Project

![hero image](https://github.com/HsuChe/VBA_challenge/blob/8c86f907f10c59b6e711a9ca5ea35a34658f35d4/Images/hero_image.jpg.jpg)

There is nothing more important to successful investment than to explore key business indicators and gauge growth for stocks. In this homework we are going to create a summary table for stock specifically from 2014, 2015, and 2016.

Features of the dataset:
* The dataset is divided primarily between three sheets for each of the years that are being analyzed, starting with 2014 and ending on 2016.
* The following are columns provided by the dataset: 
    * ticker name: **Name of the stock**
    * date: **Date the data originates**
    * opening price: **Price the stock opened for the date**
    * highest price: **The highest price the stock achieve that day** 
    * lowest price: **The lowest price the stock achieve that day**
    * closing price: **The final price before the market closes for the day**
    * trade volume: **The volume that the stock was traded for that day**
* The dataset is organized in alphabetical order where the same ticker name are listed one after another. 

* Download Dataset click [HERE](https://github.com/HsuChe/VBA_challenge/blob/859645443db611a216dec442c8de9bc2721df457/Resources/Multiple_year_stock_data.xlsx)

The homework is interested generating a few specific items for the summary table.

* The unique ticker names from the dataset
* The price change year on year
* The percentage price change year on year compared to opening price
* The total traded volume for the ticker for a given year based on the sheets.

## Special Notes: 

* I will be using Range() over Cells() in my calculations because it is easier for me to label a column as shown on excel rather than counting the columns each time
* I will be iterating up for the conditional because I can set less values to ClosingPrice/OpeningPrice/TickerName if it is done this way.

<!-- GETTING STARTED -->
## Obtaining list of unique ticker names and total volume traded.

To generate total volume trade, we have to create a volume counter that iterates through each row. 

* For loop on column A
  ```sh
  For i in 2 to LastRowA:
    TotalVolume = TotalVolume + Range("G" & i)
  Next i
  ```

To generate unique names we iterated through each of the rows comparing the value of the current cell in column A with the next iteration, or index + 1. Also make sure that the loop starts at i = 2 as the first index is the header

* For loop on column A
  ```sh
  For i in 2 to LastRowA:
    TotalVolume = TotalVolume + Range("G" & i)
    If Range("A" & i) <> Range("A" & i + 1) Then
        TickerName = Range("A" & i)
  Next i
  ```

### Prerequisites

Some important variables we need to define for the For loop are LastRowA, which will determine the last used cell in the column and set the range for our loop. The method I used is the following. 

<a href="https://stackoverflow.com/questions/39470412/last-row-in-column-vba"><strong>Code Credit »</strong></a>

* LastRowA
  ```sh
  LastRowA = Worksheet.Cells(Rows.Count, "A").End(xlUp).Row
  ```

We also need to generate the needed headers for the columns that our calculations will return to.

* LastRowA
  ```sh
  Range("I1") = "<ticker>"
  Range("J1") = "<yearly change>"
  Range("K1") = "<percent change>"
  Range("L1") = "<total volume>"
  ```

## Obtaining features needed to calculate price change year on year.

To calculate price change year on year, we would need the first opening price and last closing price for each ticker. YearlyChange = ClosingPrice - OpeningPrice

When the conditional for unique ticker ID triggers, we can get the closing price from that index as well.

* Obtain closing price from the index (i)
  ```sh
  ClosingPrice = Range("F" & i)
  ```

We can extract the opening price for the next unique ticker name by extracting the opening price of i + 1. Because we are extracting opening price for the next iteration of unique ticker name, we would have to declare the first instance of opening price outside of the for loop.

* Obtaining Closing Price from our index
  ```sh
  Dim OpeningPrice as Double
  OpeningPrice = Range("C1")

  For i in 2 to LastRowA:
    TotalVolume = TotalVolume + Range("G" & i)
    If Range("A" & i) <> Range("A" & i + 1) Then
        TickerName = Range("A" & i)
  ...
  ```

The fact that our opening price will be extracted after the current iteration of unique value is calculated, we would have to calculate the price change year on year as well as the percentage change before updating a new opening price from the conditional.

* Store Yearly Change and Yearly Percent Change to memory
  ```sh
  YearlyChange = ClosingPrice - OpeningPrice
  ```

## Error checking for Opening Price to calculate percentage change year on year.
If yearly percent change uses an opening price that is 0, will draw an error for our calculation. We must set exceptions for when opening price is 0 on the first iteration of a unique ticker.

To do this, we would need to insert two different error checkers: 

The first error check is within the conditional when a new unique value is found and before the new opening price is updated. This error will only occur if an entire ticker name does not contain an opening price value from Jan 1 to Dec 31. 

* If OpeningPrice is 0 through every iteration for a ticker, then YearlyPercent is 0
  ``` sd
  If OpeningPrice = 0 Then
    YearlyPercent = 0
  Else
    YearlyPercent = YearlyChange / OpeningPrice
  ```

The Second will be outside the for loop as error checking needs to happen on a per row basis and update opening price to the next none 0 value when it can.

* Second error checking outside the conditional that triggers when unique value is found
  ```sh
  If OpeningPrice = 0 Then
    OpeningPrice = Range("C" & i)
  ```

After the calculation is done, we can go ahead and update the opening price for the next unique value.

* Extract Opening Price
  ```sh
  OpeningPrice = Range("C" & i + 1)
  ```

## Returning all values to the correct column

We can now return the values we calculated from the if statement based on an index specifically designated  for our summary table.

* Return Values to their rightful index in the summary table
  ```sh
  Range("I" & UniqueCounter) = TickerName
  Range("J" & UniqueCounter) = YearlyChange
  Range("K" & UniqueCounter) = YearlyPercent
  Range("L" & UniqueCounter) = TotalVolume
  ' format percent '
  Range("K" & UniqueIndex).NumberFormat = "0.00%"
  ```

After TotalVolume is returned to the right column, it needs to be resetted for the next unique value and we can update UniqueIndex to make sure the next update on our summary table is in the right cell.

```sh
  TotalVolume = 0
  UniqueIndex = UniqueIndex + 1
```

## Color code positive change and negative percent change on the summary table.

The next step is to color code the positive yearly changes green and negative changes red.

* The ColorIndex for Green is 4 and The ColorIndex for Red is 3
``` sh
  If YearlyPercent > 0 Then
    Range("J" & UniqueIndex).Interior.ColorIndex = 4
  Else
    Range("J" & UniqueIndex).Interior.ColorIndex = 3
  End If
```

## Here are the results we obtained for each year starting with 2014 and ending in 2016

The summary is divided to be the first page and the rest of the pages 

* 2014 Summary page 1, get rest of the pages at [link](https://github.com/HsuChe/VBA_challenge/blob/e12b743317d4bd4271f2a3646d03661e0e8842d7/Images/2014_summary.pdf)
* 2015 Summary page 1, get rest of the pages at [link](https://github.com/HsuChe/VBA_challenge/blob/e12b743317d4bd4271f2a3646d03661e0e8842d7/Images/2015_summary.pdf)
* 2016 Summary page 1, get rest of the pages at [link](https://github.com/HsuChe/VBA_challenge/blob/e12b743317d4bd4271f2a3646d03661e0e8842d7/Images/2016_summary.pdf)


## Bonus Questions

The bonus questions asked to find 3 additional calculations. These calculations are the ticker name of the greatest percent increase and the percentage value, the ticker name of the greatest percent decrease and the percentage value, and ticker name of the largest total traded volume in the year and the traded value.

To find out the greatest percentage and the least precentage and greatest total traded we came up with 6 new varialbes.

* We are storing the value and ticker name of the parameters
  ``` sh
  Dim GreatestValue As Double
  Dim GreatestPercent As Double
  Dim GreatestDecrease As Double
  Dim GreatestDecreaseTicker As String
  Dim GreatestPercentTicker As String
  Dim GreatestVolumeTicker As String
  ```


* Set the variable values to 0
  ``` sh
  GreatestValue = 0
  GreatestPercent = 0
  GreatestDecrease = 0
  ```

We created a for loop for the summary table and compared our variables cell by cell and updating each variable appropriately.

* We are storing the value and ticker name of the parameters
  ``` sh
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
  ```

After we extracted the max for percent change and total traded volume and min for percent change, we can return them to the bonus summary table

* We are storing the value and ticker name of the parameters
  ``` sh
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
  ```

We can now return the extracted data to the correct column and rows of the bonus summary table.

* return the values to the bonus summary table
  ``` sh
  Range("O2") = GreatestPercentTicker
  Range("O3") = GreatestDecreaseTicker
  Range("O4") = GreatestVolumeTicker
  Range("P2") = GreatestPercent
  Range("P3") = GreatestDecrease
  Range("P4") = GreatestVolume
  ```

## Looping through all the worksheets

Now we can create a for loop for each worksheet in the workbook and generate both summary tables for each sheet.

``` sh 
For Each Worksheet In Worksheets
  Worksheet.Activate

  **Macro**

Next Worksheet
```

The worksheet cycle will have to be at the beginning of the macro and it would need to reset all the variables for each worksheet to start anew. 

``` sh
Ticker = ""
YearlyChange = 0
YearlyPercent = 0
TickerVolume = 0
ClosingPrice = 0
UniqueIndex = 2
```

# Conclusion

This VBA homework assignment allows us to iterate through over 2 million rows of stocks and perform analysis on these information. This demonstrates the power of VBA to perform massive tasks while maintaining all the trappings of Microsoft Excel's GUI. It really is the best of both words. 
