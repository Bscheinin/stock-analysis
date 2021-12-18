# An Analysis of Stock Performance using Refactored VBA Code

## Overview of the Project
This project used Visual Basic for Applications to analyze stock market data to assist investors in making wise decisions. 

### Purpose
The purpose of this challenge was to take previously used VBA code and to measure the performance improvement of the code using time measures. Analysis of the data included determining stock performance for two different years using VBA formulas and formatting to make the results easily digestable.

### Analysis of Code Performance
The results of the code performance analysis shows that using nested 'for loops' created a longer run time. This shows my code with a nested loop:
![Slow code](https://github.com/Bscheinin/stock-analsys/blob/main/Resources/Slow%code.png)
Total run time for this code was about 0.9 seconds. Here are each of the run time message boxes for the years 2017 and 2018.
![Slow code timing 2017](https://github.com/Bscheinin/stock-analysis/blob/main/Resources/Slow%code%timing%2017.png)
![Slow code timing 2018](https://github.com/Bscheinin/stock-analysis/blob/main/Resources/Slow%code%timing%2018.png)

By refactoring the code and changing the nested loop into two separate iterations, the run time was reduced to 0.1 seconds. Here is my code with the run time message boxes. 
![Fast code](https://github.com/Bscheinin/stock-analsys/blob/main/Resources/Fast%code.png)
![Fast code timing 2017](https://github.com/Bscheinin/stock-analysis/blob/main/Resources/Fast%code%timing%2017.png)
![Fast code timing 2018](https://github.com/Bscheinin/stock-analysis/blob/main/Resources/Fast%code%timing%2018.png)
The reason for the reduction of the run time is the first set of code with the nested loop required the codes to loop through the 3,013 rows of data twelve times. The improved code initialized tickerVolumes for the 12 rows first and then a second loop ran through the 3,013 rows just once to return the desired information.  

# Summary
There are advantages and disadvantages to refactoring code. The advantages of refactoring code is the ability to clean and improve code over time. Refactoring code can improve its performance and utility. Once you have improved code, you may be able to reuse that code in future projects which could save overall project time. Finally, a bonus to refactoring code is the collaborative effort. The old saying "Two heads are better than one" applies to VBA code as well. The overall code once refactored should be easier to understand and cleaner or the final solution is more elegant than if just one person worked on it.

Refactoring code can have disadvantages mainly in the area of time. When refactoring another person's code, it may take time to understand the flow of the code especially if it is not well documented. Once the intent is understood, it may take time to change and test those changes. If you are planning on refactoring your own code, you may not be diligent in thinking through the flow of your code as you know you will revisit it. This may be more time consuming than writing good code from the beginning. 

In refactoring this project's orginal VBA script the advantages outweighed the disadvantages. This code was well documented so seeing where improvements could be made was evident and the end result was better code performance.
