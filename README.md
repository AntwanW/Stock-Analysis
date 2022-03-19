# stock-analysis

## Overview of Project

### Purpose

This analysis aimed to edit or refactor previous code in Excel VBA to compile stock information for 2017 and 2018 to determine if the stock is worth investing in. The challenge was to increase the efficiency of the original code and determine if it was successful. Then explain my findings. 

### Results

The goal is to retrieve the ticker, the total daily volume, and the return on each stock from two data charts of stock information. The stock information incorporates a ticker value, the date the stock was issued, the opening and closing, the highest and lowest price, the adjusted closing price, and the volume of the stock. The information provided outlined steps to help refactor the original code. The steps were then listed out to set the structure for the refactoring. The screenshots of the code are listed below.


        '1a) Create a ticker Index                 tickerIndex = 0        '1b) Create three output arrays                    'The tickerVolumes array should be a Long data type.            Dim tickerVolumes(12) As Long                        'The tickerStartingPrices and tickerEndingPrices arrays should be a Single data type.            Dim tickerStartingPrices(12) As Single            Dim tickerEndingPrices(12) As Single            ''2a) Create a for loop to initialize the tickerVolumes to zero.            For i = 0 To 11                            tickerVolumes(i) = 0            Next i        ''2b) Loop over all the rows in the spreadsheet.                        For i = 2 To RowCount        '3a) Increase volume for current ticker                       tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value                '3b) Check if the current row is the first row with the selected tickerIndex.        'If  Then                    If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then                    tickerStartingPrices(tickerIndex) = Cells(i, 6).Value                                End If                '3c) check if the current row is the last row with the selected ticker        'If  Then                    If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then                    tickerEndingPrices(tickerIndex) = Cells(i, 6).Value                                    End If                        '3d Increase the tickerIndex.                    If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then                    tickerIndex = tickerIndex + 1                                    End IfNext i'4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.            For i = 0 To 11                Worksheets("All Stocks Analysis").Activate            Cells(4 + i, 1).Value = tickers(i)            Cells(4 + i, 2).Value = tickerVolumes(i)            Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1



### Summary

#### What are the advantages or disadvantages of refactoring code?

* The advantage of refactoring is that the code is less complex, organized, and easier to understand and read. It could help with debugging and faster programming. The disadvantage is that you might have to retest lots of functions, which is time-consuming.

#### How do these pros and cons apply to refactoring the original VBA script?

* The advantage is that refactoring could have better quality. The disadvantage is not knowing where to go next after changing a line and trying to debug. Unsure if refactoring helps with run time or if your coding is correct.

![This is an image](https://github.com/Wrancher123/stock-analysis/blob/main/Resources/VBA_Challenge_2017.png)

![This is an image](https://github.com/Wrancher123/stock-analysis/blob/main/Resources/VBA_Challenge_2018.png%20.png)


