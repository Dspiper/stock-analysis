# Stocks Analysis

## Overview of Project 
### Purpose
In this project, we wanted to help Steve by expanding the dataset to handle the entire stock market over the last few years instead of a dozen stocks. We refactored the code we previously wrote to be more efficient and handle the larger dataset, so Steve will be able to do more research for his parents. Link to the dataset can be found here: [VBA_Challenge](https://github.com/Dspiper/stock-analysis/blob/main/VBA_Challenge.xlsm)

## Results

Comparing the how the stocks performed in 2017 compared to 2018, we see that 2017 was an overll better year for returns. The only stock that had a negative return was TERP, which had a -7.2% return with a total daily volume of 139,402,800. The stock with the highest return in 2017 was DQ with a return of 199.4%, and the stock with the highest daily volume was SPWR at 782,187,000. 

In 2018, we saw a negative return on all stocks except RUN and ENPH. RUN had a return of 84% while ENPH had a slightly lower return at 81.9%. ENPH also had the highest daily volume 607,473,500. Unlike in 2017, DQ had the worst return in 2018 at -62.6%. 

Before refactoring the code, we were seeing a run time for 2017 and 2018 of 0.2773438 seconds. After refactoring the code, we saw a signifacant improvement in the run time of the code. The refactored code for 2017 ran in 0.0546875 seconds, and the refactored code for 2018 ran in 0.05859375 seconds. Please see images below. 

![VBA_Challenge_2017](https://github.com/Dspiper/stock-analysis/blob/main/Resources/VBA_Challenge_2017.png) 
![VBA_Challenge_2018](https://github.com/Dspiper/stock-analysis/blob/main/Resources/VBA_Challenge_2018.png)

Some significant changes were made when refactoring the code. The biggest change was made in the if statements in the for loop. In the original code, we used the if statements `If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then` to find the starting price of the current ticker and `If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then` to find the ending price of the current ticker. In the refactored code, these if statments were modified to `If Cells(i - 1, 1).Value <> tickers(tickerIndex) Then` and `If Cells(i + 1, 1).Value <> tickers(tickerIndex) Then` to find the starting and ending and check if the current row is the first or last using the ticker index. We removed the And statement from the refactored because it is redundent. Also, we created the output arrays to hold the Ticker, Total Daily Volume, and Return.

```
Dim tickerVolumes(12) As Long
Dim tickerStartingPrices(12) As Single
Dim tickerEndingPrices(12) As Single
```

## Summary

In a summary statement, address the following questions.
What are the advantages or disadvantages of refactoring code?
How do these pros and cons apply to refactoring the original VBA script?
