# VBA_Challenge

## Overview of Project
The purpose of this analysis is to refactor code for Steve's research on the stock market over the last few years, to make the VBA Script for the green_stocks spreadsheet run faster, by making the code more efficient. 

## Results
Refactoring the code did make it run a little faster. Before I refactored it, the code took about 0.53 seconds to run. After refactoring it and running it a few times, it took about 0.3 seconds to run. 

![All_Stocks_Analysis_Refactored_2017](https://github.com/sjwedlund/VBA_Challenge/blob/main/All_Stocks_Analysis_Refactored_2017.png?raw=true) 
![All_Stocks_Analysis_Refactored_2018](https://github.com/sjwedlund/VBA_Challenge/blob/main/All_Stocks_Analysis_Refactored_2018.png)

As you can see from the screen shots of my Excel sheet for **All Stocks 2017** (refactored), all stocks except *TERP* were "in the green", that is- had a positive return. *TERP* was "in the red" with a return of -7.2%. Comparitively, there were only two stocks with a positive return for the year 2018, which were *ENPH* and *RUN*, at 81.9% and 84.0%, respectively. All other stocks for 2018 had a negative return. 

# Run Time
![green_stocks_2017](https://github.com/sjwedlund/VBA_Challenge/blob/main/resources/resources/green_stocks_2017.png)

Here in this screen shot of the run time for my original macro for All Stocks Analysis, you can see that the code ran in 0.53125 seconds for the year 2017. 

![VBA_Challenge_2017](https://github.com/sjwedlund/VBA_Challenge/blob/main/resources/VBA_Challenge_2017.png)

In the above image, a screen shot of the popup for the refactored code, you can see that the run time for 2017 was a bit faster, at only 0.3125 seconds. 

