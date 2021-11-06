# 2nd week challenge
## The purpose of this analysis 

We are running the analysis to support the client stock investment decision. 
So we run the VBA script and format the result to get an easy view of the return on the investment. 
For each stock, we have its annual volume and the return rate in %, and we mark the stocks that are profitable in the green color and red for the opposite.

## Results:  
Looking at the 2017 and 2018 data, the overall volume is decreasing. 
The return is declining. In 2017, 11 stocks have a positive return, but in 2018, it was down to only 2. 
![VBA_Challenge_2017](https://github.com/Sirius0531/stocks-analysis/blob/main/Resources/VBA_Challenge_2017.png)
![VBA_Challenge_2018](https://github.com/Sirius0531/stocks-analysis/blob/main/Resources/VBA_Challenge_2018.png)
The suggestion for the client is to either relocate their investment to different stocks or other industries.  

## Summary: 
**What are the advantages or disadvantages of refactoring code?**
By using an array, a uniform array name and different subscripts can be used to determine the only element in the array. 
Significantly shorten and simplify the program code, thereby improving the efficiency of the application. The time for process the analysis went down from 0.5 seconds to 0.125 seconds. 
![VBA_Challenge_2018](https://github.com/Sirius0531/stocks-analysis/blob/main/Resources/VBA_Challenge_Asssigned_2018.png)
This is for 3000 lines of data, it will have a bigger difference when we analyze a larger data set.  
**How do these pros and cons apply to refactoring the original VBA script?**
To apply refactoring script, make the analysis more efficient. But the original script, with 2 variables(i&j) and For-loop is easier for beginners to understand the logic.
The refactoring version of the code is more complecated. We need to define the range for the arrays, in this case, we need to let the code know we are running the loop for 12 stocks. Second is to mind for assing the variable to the for-loop, in this case, the tickerIndex.
