# Stock Analysis
<img src="Resources_Mod2/2017_2018_Analyses.png" width="300" height="150">

## Table of Contents
* [Overview of Project](https://github.com/rkaysen63/stock-analysis/blob/master/README.md#overview-of-project)
* [Results](https://github.com/rkaysen63/stock-analysis/blob/master/README.md#results)
* [Summary](https://github.com/rkaysen63/stock-analysis/blob/master/README.md#summary)

## Overview of Project
The customer's parents had solely invested in DAQO New Energy Corp, a company that makes silicon wafers for solar panels.  The customer believes that his parents should diversify their investments and has asked for help analyzing 2017 and 2018 stock data in order to evaluate DAQO (ticker symbol DQ) and other green energy stocks.  Using VBA macros, the Total Daily Volume and Return were calculated for each stock for the years 2017 and 2018.

For ease of use, buttons were provided to clear the page and then run the analysis with a prompt for the year.  

The macros were coded twice, in different ways, in order to demonstrate refactoring and the difference in run time.

(Data set and premise from BootCamp Module 2 Challenge: https://courses.bootcampspot.com/courses)

## Results

### Analysis
Over three thousand rows of data per year were compiled.  Total Daily Volume and Return were calculated for each ticker for each year.  The output was formatted to show gains in green and losses in red.

The side by side comparison below shows the results from two years of data.  DQ performed well in 2017 but suffered huge losses in 2018.  ENPH and RUN were the only two stocks with positive returns for both years with RUN making the most gains. 

![alt text](Resources_Mod2/2017_2018_Analyses.png)

https://www.mrexcel.com/board/threads/how-do-you-align-text-center-in-a-cell-using-vba.276160; https://docs.microsoft.com/en-us/office/vba/api/excel

### Nested Loops
To obtain the same final results, two macros were coded in different ways.  The first macro *StockAnalysis* used nested loops.  The outer loop initiated a particular ticker, set the total volume to zero and launched into the inner loop.  The inner loop circulated through a particular ticker in order to calculate Total Volume for that ticker and obtain its Starting and Ending Price.  Then the output cells were filled before moving on to the next ticker. At the next ticker, the process was repeated.

   `For i = 0 To 11; ticker = tickers(i); totalVolume = 0`
       
       For j = 2 To RowCount
        If Cells(j, 1).Value = ticker Then
          totalVolume = totalVolume + Cells(j, 8).Value
        End If
        If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
          startingPrice = Cells(j, 6).Value   
        End If
        If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
          endingPrice = Cells(j, 6).Value
        End If
       Next j
        
### Arrays
When refactored, the data in *StockAnalysis Refactored* was compiled using multiple arrays rather than nested loops.  First a loop was created to loop through the tickerVolumes and set them to zero.  A second loop was created to cycle through the arrays to calculate Ticker Volume and obtain the Starting and Ending Prices for each ticker by using a new variable called tickerIndex.  This loop ended by moving to the next ticker.  A final loop, populated the output. 
    
  `For i = 0 To 11; tickerVolumes(i) = 0; Next i`
   
  'For i = 2 To RowCount; tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value`
        
        If Cells(i - 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
          tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
        End If`

        If Cells(i + 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
            tickerIndex = tickerIndex + 1
        End If
    Next i`
    
### Comparison of Run Times - Original Code vs Refactored Code

<img src="Resources_Mod2/AllStocks2017.png" width="430" height="250">  <img src="Resources_Mod2/AllStocksRefactored2017.png" width="430" height="250">
<img src="Resources_Mod2/AllStocks2018.png" width="430" height="250">  <img src="Resources_Mod2/AllStocksRefactored2018.png" width="430" height="250">

## Summary
* In general, refactoring has the advantage of optimizing code and possibly reducing run time.  The obvious disadvantage is that it requires additional thought and time that actually lead to another advantage, which is learning new code.
* For this particular exercise, refactoring cut the run times significantly although I do not trust the run time numbers for the original code.  It only took seconds to run the original code, but run time shows 57300 secs = 14.5 minutes. I learned a different way tackle the problem, but the disadvantage is that I spent a considerable amount of time trying to understand and solve the problem.  Unlike the nested lookes, the method using the arrays was not as intuitive to me.  
