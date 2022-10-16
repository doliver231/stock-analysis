# VBA of Wall Street

## Overview of Project

Steve, a recent finance graduate, wants to apply what he has learned to help his parents diversify their stock funds. A list of several green energy stocks are analyzed with the use of Excel's extension: Visual Basic for Applications (VBA).

### Purpose

Using VBA's full capability, we want to create a code to automate the analyses performed on the list of green energy stocks. We also want to refactor it in such a way that it can be applied to the entire stock market for the last few years, and ran in a practical and timely manner.

## Analysis

Using a data set from a total of 12 different ticker stocks, a VBA script was created to calculate the Total Daily Volume and the Yearly Return for years 2017 and 2018. The macro created involved an extensive code, which included variables, arrays, for & nested loops, and conditional statements. In hopes to reduce the execution time of the script and make it functional for potential larger data sets, the code was refactored, involving a total of four arrays and concise conditional statements. Buttons and formatting were introduced to make the worksheet more interactive and user-friendly. The refactored VBA macro has been attached.

### Refactored Code

![Refactored VBA Code - Part 1](https://github.com/doliver231/stock-analysis/blob/main/Resources/Refactored_Code_1.png)
![Refactored VBA Code - Part 2](https://github.com/doliver231/stock-analysis/blob/main/Resources/Refactored_Code_2.png)
![Refactored VBA Code - Part 3](https://github.com/doliver231/stock-analysis/blob/main/Resources/Refactored_Code_3.png)
![Refactored VBA Code - Part 4](https://github.com/doliver231/stock-analysis/blob/main/Resources/Refactored_Code_4.png)

## Results

There are notable differences between the results for 2017 and 2018. 92% of the stocks in 2017 showed positive yearly returns, while 83% of the stocks in 2018 had negative returns. 

In 2017, the one stock that stood out was TERP, with a return of -7.2%. DQ, ENPH, and SEDG had the highest returns. DQ had the highest yearly return out of all, but shows the least daily volume. In 2018, the only ones with positive returns are ENPH and RUN. ENPH even had the highest total daily volume out of all stocks. If I had to suggest a stock to Steve's parents, it would be to invest in ENPH, which showed the best progress for both years.

The execution times of the refactored VBA macros compared to the original scripts differed substantially. Without the code being refactored, the execution times for both 2017 and 2018 were 1.09 and 1.05 seconds, respectively. On the other hand, after being refactored, the times to the execute the code for 2017 and 2018 were 0.16 and 0.14 seconds, respectively. Clearly, the refactored macro produced much faster results.

### VBA Code Execution Time

##### 2017 All Stocks Analysis - Original Script
![Original Execution Time for 2017](https://github.com/doliver231/stock-analysis/blob/main/Resources/VBA_Challenge_2017_Original.png)

##### 2018 All Stocks Analysis - Original Script
![Original Execution Time for 2018](https://github.com/doliver231/stock-analysis/blob/main/Resources/VBA_Challenge_2018_Original.png)

##### 2017 All Stocks Analysis - Refactored Script
![Refactored Execution Time for 2017](https://github.com/doliver231/stock-analysis/blob/main/Resources/VBA_Challenge_2017.png)

##### 2018 All Stocks Analysis - Refactored Script
![Refactored Execution Time for 2018](https://github.com/doliver231/stock-analysis/blob/main/Resources/VBA_Challenge_2018.png)

## Summary

### Advantages and Disadvantages of Refacturing Code

An advantage to refactoring any general code is the script will be more concise, organized, and easier to follow. From what we saw in the analysis, refactored code tends to reduce the time to execute the whole code, which is ideal for huge datasets.

A disadvantage is that it can be risky to change any section of an already set code. It may introduce bugs and may be difficult to interpret for beginner coders.

### Advantages and Disadvantages of Original & Refactored VBA Scripts

I believe there are always advantages when refactoring than without. In this specific case, the refactored VBA script versus the original has clear positive results. The observable outputs were exactly the same, however, it made the code flexible and removed any needless complexity. At a small scale, the execution times were greatly reduced after refactoring, which is important when dealing with much bigger datasets.

Refactoring the original script can always be risky. In my case, I ended up debugging a few times because certain subroutines and statements were mismatched and easily gone unnoticed. With a complex code, one really has to decide if it's really worth risking breaking the code by trying to make it better. It worked for this challenge but may not be useful for data sets of other sizes and categories. The risk of complicating the code was disadvantage on its own.