# Stocks Analysis

## Overview of Project 
### Purpose
In this project, we wanted to help Steve by expanding the dataset to handle the entire stock market over the last few years instead of a dozen stocks. We refactored the code we previously wrote to be more efficient and handle larger datasets, so Steve will be able to do more research for his parents. Link to the workbook can be found here: [VBA_Challenge](https://github.com/Dspiper/stock-analysis/blob/main/VBA_Challenge.xlsm)

## Results

Comparing how the stocks performed in 2017 to 2018, we see that 2017 was an overall better year for returns. The only stock that had a negative return was TERP, which had a -7.2% return with a total daily volume of 139,402,800. The stock with the highest return in 2017 was DQ with a return of 199.4%, and the stock with the highest daily volume in 2017 was SPWR at 782,187,000. Results for all stocks in 2017 can be found below. 

![VBA_Challenge_2017](https://github.com/Dspiper/stock-analysis/blob/main/Resources/VBA_Challenge_2017.png) 

In 2018, we saw a negative return on all stocks except ENPH and RUN. RUN had a return of 84.0% while ENPH had a slightly lower return at 81.9%. ENPH also had the highest daily volume in 2018 at 607,473,500. Unlike in 2017, DQ had the worst return of all stocks in 2018 at -62.6%. Results for all stocks in 2018 can be found below.

![VBA_Challenge_2018](https://github.com/Dspiper/stock-analysis/blob/main/Resources/VBA_Challenge_2018.png)

Before refactoring the code, we were seeing a run time for the code of 0.2773438 seconds for both 2017 and 2018. After refactoring the code, we saw a significant improvement in the run time of the code. The refactored code for 2017 ran in 0.04296875 seconds, and the refactored code for 2018 ran in 0.046875 seconds. The run times for the refactored code can be found in the above images with the 2017 and 2018 stock results. 

Refactored code can be found here: [VBA_Challenge](https://github.com/Dspiper/stock-analysis/blob/main/VBA_Challenge.vbs)

## Summary
What are the advantages or disadvantages of refactoring code? An advantage of refactoring code is that it allows you to make your code more efficient and easier for you to understand and use. A disadvantage of refactoring code is that not everyone may refactor the code the same way. Two people may achieve the same end result, but may have refactored their code differently which could be confusing to someone trying to understand and use the code in the future.   

How do these pros and cons apply to refactoring the original VBA script? When refactoring the original VBA script, I found that several areas of the original code were redundant which was causing my run time to increase. By removing the redundant parts of the code, it was much more efficient and easier to use and understand. By going to the office hours, I got to see examples of how my classmates were factoring their code. This gave me insight into how the code can be approached from different angles to reach the same end result. 
