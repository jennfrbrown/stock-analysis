# stock-analysis
Module 2 VBA

## Overview of Project

### Purpose
The purpose of this project was to refactor the Module 2 solution code by switching the nesting order of the "for" loops.  We would then determine if the VBA script ran faster due to the nesting order of the "for" loop change.

### Background
During Module 2 we built a workbook that analyzed a dataset that focused specifically on green energy stocks (for the years 2017 and 2018) to determine their Yearly Return Percentage and Total Daily Volumes.  Within the green energy stock category, our client  was particularly looking at DAQO stocks (Ticker: DQ).  The analysis helped determine if DQ's performance over the years compared to other green energy stocks made them a good investment.  

Our client now want to expand the data set to include the entire stock market over the last few years.  Refactoring the code was important as the original workbook was originally coded to analyze only a dozen or so stocks, not thousands.  To perform this ananalysis, I used Microsoft Visual Basic (VBA) incorporating conditional formatting, loops, and conditional statements.  As the purpose of this project focuses on refactoring code, the results will highlight the differences between the two coding orders and format.

## Results
Each table below shows the Ticker Name, Total Daily Volume, and Percentage Return for that specific year.
<img src="https://github.com/jennfrbrown/stock-analysis/blob/master/Images%20for%20ReadMe/2017%262018.png" height = 300>

### Performance
All green stocks analyzed in 2017 had positive yearly returns.  However, in 2018, all but two stocks had negative yearly returns. In 2017, DQ had low total Volume, yet a high rate of return.  In 2018, DQ had a higher total volume, but a lower rater of return.  However, all but one of the twelve stocks analyzed had a decrease in yearly return.

### Code
The code for both "All Stocks Analysis" and "All Stocks Analysis Refactored" produce the same results.  In refactoring the code, it shows that even though both programs produce the same output, their may be faster, more efficient ways to code a program.

The major difference between the two codes is that "All Stock Analysis" ran the analysis for a specific ticker on a worksheet holding the data, printed the output on an output worksheet, and then went back to the data worksheet to run the analysis for the next ticker.  The way the program was coded meant that it was constantly switching back and forth between worksheets.

The code for "All Stocks Analysis Refactored" ran the analysis by gathering all the data in the data worksheet, storing it in an array, and then printing all outputs in the output worksheet.  This means there were was no need to constantly switch back and forth between worksheets, which produced a faster program run time.




## Summary
