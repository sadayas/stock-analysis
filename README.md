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

Instead we were able to reuse a different code from stackoverflow, which processed through all data within the rows instead.

```
RowCount = Cells(Rows.Count, "A").End(xlUp).Row
For i = 2 To RowCount
````

## Summary
## Pros and Cons of Refactor
There are many positives to refracting a code.   It can be updated to become a simpler and cleaner code.   The code then will be easier for others to understand, especially with less lines of code to read.  In a team setting, it will be easier to collaborate as the code will be more simplified. With a simpler code, it will lead to easier and quicker updates. Thus this will lead to saving time, and possibly money.  With a refactored code, we are able to be more efficient for our clients, and can process data quicker.  

However, there are also disvantages to code refractoring. One would need to understand the original code before even attempting to refactor.  This would be more time consuming over time or if there is a time crunch.  Another issue may involve using precision, because with some codes processing large data information at once, certain parts of the data may be missed.  While it is a quicker approach, additional lines of code would be needed again to find precise data.

In this stock analysis, the refractor code did help in certain areas.  For one, processing through the data was definitely speedy compared to the original code.  Certain parts of the refractored code did make the code simpler and more stream-lined.  The code then became quite efficient.  There are difficulties as well, the creating the refactored code was difficult.  If one did not understand how to use VBA, it would take more time than just using the older code.  There was more research needed to help develop the code than even just running it.  There was more testing needed, and if there was a due date, it would be more stressful.
