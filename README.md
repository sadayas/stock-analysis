# Stock Analysis with VBA

## Overview of Project
### Purpose
The client, Steve, wants to know how the entire stock mark was over the last few years.  With his research and data he collected for his parents, we are able to analyze and develop code to read through the data.

## Results

In our results, we are able to see that the return between 2017 is much higher than 2018. With nearly all the stocks having a positive return.  However in 2018, as shown in the images, the stocks took a negative return.  The tickers ENPH and RUN, however saw its return raise exponentially.

![This is an image](https://github.com/sadayas/stock-analysis/blob/main/VBA_Challenge_2017.png)   ![This is an image](https://github.com/sadayas/stock-analysis/blob/main/VBA_Challenge_2018.png)

With the refractored script, we were able to process the data quicker.  It took away seconds compared to the original code.  For instance, with the use of Array output, we were able to save seconds off of the analyzing time.  It was able to process the data quicker, as instead of searching for data in each cell, it was able to scan an entire index of information.  In this code, we create array output for tickerVolumes, tickerStartingPrice, and tickerEndingPrice.

```
'1b) Create three output arrays
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
```
Another instance of saving processing time, involved looping through the spreadsheet.  In our original code, we would loop through the beginning of the row, and the end of the rows.

```
rowStart = 2
rowEnd = 3013
For i = rowStart to rowEnd
Next i
````

Instead we were able to reuse a different code from stackoverflow, which processed through all the rows instead.

```
RowCount = Cells(Rows.Count, "A").End(xlUp).Row
For i = 2 To RowCount
````

## Summary
