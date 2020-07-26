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

<img src="https://github.com/jennfrbrown/stock-analysis/blob/master/Images%20for%20ReadMe/2017TimeComparison.png">

The refactored code ran approximately 4.38 times faster than the original code.  One thing to note is that as we continue to run the program, the run time will decrease slightly from its original run time.

The decrease in run time comes from removing the nested loops and instead creating arrays to hold information.  Below is a comparison of the two codes showing the differences in the order of the steps coded. (Click on picture to enlarge)

<img src="https://github.com/jennfrbrown/stock-analysis/blob/master/Images%20for%20ReadMe/CodeComparison.png">

## Summary
### Advantages and Disadvantages of Refactoring Code
The purpose of refactoring code is to improve the structure and design of the code while keeping its functionality.

#### Advantages
- Code Reability & Reduced Complexity: creating a cleaner, simpler architecture makes it easier to fix bugs
- Improved Performance: runs faster or uses less memory
- Expandibility: it is easier to expand on current capabilities if the coding uses easily identifiable patterns.

#### Disadvantages
- Effort: refactoring requires that you have some knowledge of the existing design, depending on where you enter this process it may take additional time to understand the original design/motivation
- Introduction of bugs: by refactoring the code you may be introducing bugs that were not originally there, thus increasing the effort and time commitment required to produce a better application

### Advantages and Disadvantages of the original and refactored VBA script
As mentioned above one of the major advantages of refactoring this VBA script was the decrease in runtime.  By creating an array, we also created a cleaner  code structure.

I don't know that I would identify this part of refactoring as a disadvantage, but rather a challenge.  In refactoring, you must have a firm knowledge of the original code.  If you don't know what the original code is doing, you can't improve on it.  If you don't have a firm knowledge of code in general,  you can't look at it and identify an alternate/better way to do it.



