# Green Stock Analysis 2017-18

## Project Overview:

![](https://github.com/DevTrav/stock-analysis/blob/main/resources/market.png)

The initial VB application was built in workbook green_stocks.xlsm to perform stock market performance analysis of 12 "green" stocks. The tool was refactored in VBA_Challenge.xlsm to serve as a proof-of-concept that the application may be used to perform the necessary analysis across larger data sets, at a reduced run time.

## Results

The original application ran at ~ .5 seconds when analyzing the performance of 12 stock. This could be cumbersome when using the tool on a large data set. For example, if the data set were 4,000 stock, the original code base would run  ~~ 33 min. 
![](https://github.com/DevTrav/stock-analysis/blob/main/resources/green_stocks_2017.png)
![](https://github.com/DevTrav/stock-analysis/blob/main/resources/green_stocks_2018.png)

### The refactor 

+ The refactor replaced nested conditionals with `For Loops`. This optimized the compile time significantly. Instead, we created `Arrays` to store the tickerVolumes, tickerStaringPrices and tickerEndingPricesindex. Then, we used tickerIndex to capture a count of stock increase/ decrease during 2017-18. 

+ The change from nested conditionals to looping through Arrays is more dynamic and makes the application scalable.

+ In the example below, we initialize the tickerVolume to 0 and use the tickerIndex to loop through the spreadsheet to increase the volume at the current ticker (i).

```
 '1a) Create a ticker Index
    Dim tickerIndex As Integer
    tickerIndex = 0
 


    '1b) Create three output arrays
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    For i = 0 To 11
    tickerVolumes(i) = 0
    Next i
    ''2b) Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
```

## Summary 
+ The removal the unnecessary conditional text and the implementation of the `For Loops` reduced the compile time to ~ .2 seconds for 2017 data and ~~.07 seconds for 2018 data.

![2017](https://github.com/DevTrav/stock-analysis/blob/main/resources/VBA_Challenge_2017.png)
![2018](https://github.com/DevTrav/stock-analysis/blob/main/resources/VBA_Challenge_2018.png)

+  The refactored code will offer greater clarity to the fucntion of the application and therefore easier to maintain.




