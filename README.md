# Stock Analysis with Excel VBA
## Overview of Project

### Purpose
The purpose of this project was to refactor Microsoft Excel VBA code. We took collected stock information on specific stocks in the year 2017 and 2018 and analyzed the data to determine whether or not the stocks were worth investing. This was completed in multi-step processes, with the ultimate goal being able to refactor the original code to run quicker and more efficiently.

### The Data
The data was presented in two Excel spreadsheets featuring stock information on 12 different stocks. The information included a ticker value, the date the stock was issued, the opening, closing and adjusted closing price, the highest and lowest price, and the volume of the stock. The goal was to retrieve the ticker, the total daily volume, and the return on each stock.

## Results
### Analysis
My first step for refactoring the code included copying the code needed to create the input box, chart headers, ticker array, and to activate the appropriate worksheet. The steps were then listed out in order to set the structure for the refactoring. Below is the code with the modified steps:
    '1a) Create a ticker Index
        tickerIndex = 0
    
    '1b) Create three output arrays
        Dim tickerVolumes(12) As Long
        Dim tickerStartingPrices(12) As Single
        Dim tickerEndingPrices(12) As Single
    
    '2a) Create a for loop to initialize the tickerVolumes to zero.
    'If the next row’s ticker doesn’t match, increase the tickerIndex.
        For i = 0 To 11
   
        tickerVolumes(i) = 0
        tickerStartingPrices(i) = 0
        tickerEndingPrices(i) = 0

        Next i

    '2b) Loop over all the rows in the spreadsheet.
        For i = 2 To RowCount

    '3a) Increase volume for current ticker
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value

    '3b) Check if the current row is the first row with the selected tickerIndex.
    'If  Then
        If Cells(i - 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
          tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
        End If
    
    '3c) check if the current row is the last row with the selected ticker
    'If Then
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
          tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
        End If
    
    '3d Increase the tickerIndex.
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
          tickerIndex = tickerIndex + 1
        End If

        Next i

    '4) Loop thru your arrays to output the Ticker, Total Daily Volume, and Return.
        For i = 0 To 11

        Worksheets("All Stocks Analysis").Activate
          Cells(4 + i, 1).Value = tickers(i)
          Cells(4 + i, 2).Value = tickerVolumes(i)
          Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
    
        Next i

The next step was running the refactored code to determine if there was improved speed with the analysis. In the first image, https://github.com/SThieshen/stock-analysis/blob/main/VBA_Challenge_2018.png this was run prior to refactoring our code. The second image https://github.com/SThieshen/stock-analysis/blob/main/VBA_Challenge_2018_refactored.png is analysis for the same information, but with the code refactored. As you can see when clicking on the links, the speed did improve.

## Summary
### Pros and Cons of Refactoring Code
When code is refactored, we are able to organize it better, debug it, and make the program run faster. We are also able to simplify the code, making it easier to read for others. It can also save time and money by making the code more maintainable. Unfortunately, sometimes refactoring code costs more in the short run and we don't always have the luxury of time and money to refactor it. There are also the chance new bugs could be introduced within a program when the code is refactored.

### The Advantages of Refactoring VBA Analysis
The most noted benefit of refactoring the original VBA script was the decrease in macro run time. The original analysis took almost one second to run, but the new analysis took almost a tenth of that time to run. Running at optimal efficiency is important in all programming.
