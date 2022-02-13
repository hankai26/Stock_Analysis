# Stock Analysis and VBA Code Refactoring
## Overview of Project
This project is to analyze green stock performance and to refactor the solution code to make it more efficiently. Original solution script could take long time when expanding the datashet to analyze entire stock market over the past years, especially when analyzing thousands of stocks. It can be quite helpful to develop the code with appropriate edits and to run with less calculation, using less memory and easier to read. In this project we refactored the code in multiple ways to reduce the execution time. The comparison of different algorithms is also conducted to deepen our understanding on this analysis code, as well as improving future coding.


## Code Refactoring and Results

### 1. Code Refactored in Two Ways
The solution workbook for stock analysis during multiple years is first prepared as an original script. It searches the datasheet by rows and writes data into new sheets by cells via two for-loops. This project further develops the solution in two ways to reduce the executing time.

- A. One way to refactor code is to store the calculation data into variables, like arrays, rather than overwritting spreadsheet cells in every repeating calculation. Finally, the results would be written into the subject worksheet at one time. This solution saves time on reading and writting on worksheets. However, the number of loops and total calcutions remain same. It is expected to reduce some time, but not much.

- B. In order to make the programming more efficiently, another way to refactor code is to reduce the calculations directly and loops by improving the logical algorithms. When reviewing the dataset, we can see that the datasheet for available years is sorted. All stock tickers occur together by groups. Of course, we could easily sort the data in spreadsheet by the build-in sorting function even if it was not sorted originally.
Based on the sorted and grouped stock tickers, other conditions are added to control each ticker index moving from one by one. The code is refactored to search data by rows using only one loop to read and store data into responding variables.


### 2. Comparison Stock Performance between 2017 and 2018
The output keeps consistent among the original script and the two refactored scripts. The analysis results for 2017 and 2018 are listed in Figure 1 and Figure 2.

![2017_AnalysisTable](https://github.com/hankai26/Stock_Analysis/blob/main/Resources/VBA_Challenge_2017_AnalysisTable.PNG)

**Figure 1. Analysis for All Stocks in 2017**

![2018_AnalysisTable](https://github.com/hankai26/Stock_Analysis/blob/main/Resources/VBA_Challenge_2018_AnalysisTable.PNG)

**Figure 2. Analysis for All Stocks in 2018**

Twelve stocks are analyzed in this project. The overall stock performance in 2018 is worse than that in 2017. Ten from twelve stocks in 2018 turn out with negative year returns. The green stock market probably showed a decreasing trend in 2018. Comparatively, only one stock indicates a negative year return in 2017, which means that the market probably was quite hot throughout 2017. One stock "ENPH",probably needs further research in this field, whose performance in both year was impressive.


### 3. Comparison Execution Times between Original Script and Refactored Scripts
The original solution code works fine, but kind of slowly especially with more stock analysis. To comapre the difference in the original script and the refactored scripts, we first measure the execution time for original script. By inserting the syntax of Timer, the measurement results for both 2017 and 2018 are **1.046875** seconds and **1.121094** seconds, respectively. Two reasons can be related to the comparatively long execution time for original script. One is to reapeat visiting disk to read/write into worksheet every time. The other reason is associated to the huge amount of calculations in the two loops. 

![2017_original](https://github.com/hankai26/Stock_Analysis/blob/main/Resources/VBA_Challenge_2017_Original.PNG)
![2018_original](https://github.com/hankai26/Stock_Analysis/blob/main/Resources/VBA_Challenge_2018_Original.PNG)

The code is first refactored to **Version A**, by storing every calculation in variables, rathen than writting into cells. The execution times for 2017 and 2018 decrease to **0.7851563** and **0.7382813**, respectively. This code refactoring obviously reduce the running time but within a limited range. The saved time is on account of calculating and writing into cells at one time, which avoids repeating the operation of opening and saving worksheet.

![2017_Refactor1](https://github.com/hankai26/Stock_Analysis/blob/main/Resources/VBA_Challenge_2017_Refactor1.PNG)
![2018_Refactor1](https://github.com/hankai26/Stock_Analysis/blob/main/Resources/VBA_Challenge_2018_Refactor1.PNG)

Further, the refactoring code in **Version B** is to reduce the two loops to one, which directly reduces the number of calculations to one twelfth of that in the original script. The time complexity is reduced from **O(n^2)**(original script and refactored script A) to **O(n)**(refactored script B). Therefore the execution time is significantly decreased to **0.1210938** seconds and **0.1054688** seconds for the two year analysis.

![2017_Refactor2](https://github.com/hankai26/Stock_Analysis/blob/main/Resources/VBA_Challenge_2017_Refactor2.PNG)
![2018_Refactor2](https://github.com/hankai26/Stock_Analysis/blob/main/Resources/VBA_Challenge_2018_Refactor2.PNG)


## Summary
- Code refactoring is a process to edit the codes to achieve the same functions. To make the code more efficiently, it's critical to run the code with less calculation or less memory. Refactoring code is capable of reducing the calculation numbers and disk operations, as well as the total time to execute the code. Of course, to reduce operation numbers may require a little bit more computer memory for variables.
- In this project, the execution time for original script is comparatively long. However, it could directly alter the content in cells one by one over time and make any revision **in-place** without extra variables and memory usage. The time complexity is **O(n^2)**.
- The refactoring script A avoids frequent file opening and disk operations, by adding array variables. It then reduces the running time somehow by sacrificing a little bit memory. The time complexity doesn't change.
- The refactoring script B looks like a better version with **significantly reduced execution time**. This script works with much less calculation and CPU operation, via array variables using a little bit more memory. The time complexity is reduced to **O(n)**. Normally less calculation is much more critical when solving complicate problems. This script B is a more efficient code to solve the problem, even though it requires a well sorted and grouped datasheet prior to loop calculation.
