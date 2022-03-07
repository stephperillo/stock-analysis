
# Stock Analysis Using VBA

## Overview of Project

In this project, I was tasked with analyzing green stocks to find total daily volume and annual return. I used VBA and Excel to create subroutines in order to analyze the dataset, which can also run for either given year. I also included a timer in order to time the amount of time it takes for each script to run. Then I refactored the script in order to make the script run faster. 

## Results

To make the code more efficient, I looped through all the data one time in order to collect the information. I did this by creating 3 arrays to hold tickerVolumes, tickerStartingPrices, and tickerEndingPrices. These arrays store data for each stock when the for loop runs on each. 

###### Original 



###### Refactored 



As a result of refactoring the code, it ran 0.6210937 seconds faster for 2017 and 0.5039062 seconds faster for 2018 stocks. 
![Refactored 2017](https://github.com/stephperillo/stock-analysis/blob/main/resources/VBA_Challenge_2017.png)

![Refactored 2018](https://github.com/stephperillo/stock-analysis/blob/main/resources/VBA_Challenge_2018.png)

Reading the data via an array has proven to be faster than reading in each cell individually in a nested loop.

I confirmed that the stock analysis outputs for 2017 and 2018 are the same for the refactored code as during the original code.


## Summary

###### What are the advantages and disadvantages of refactoring code?

In general, refactoring code is important. The main advantage is that it can make the code more efficient; the new code can use less memory, take less time to run, and make it easier for others or even myself to understand or read in the future. 

Refactoring code can also have its disadvantages. There is always the possibility of messing up perfectly functioning code when trying to refactor the original. It is also possible to complicate the code or make it more difficult to understand for someone else. 

###### What are the advantages and disadvantages of the original and refactored VBA script in this project?

For this project, improving the script made it run faster. Doing so also organized it better and made it easier to comprehend. It also facilitates any possible future modifications by nesting the loops. If that time arises, it will be easier to adjust the code in one area of the code as opposed to manually changing the code in several different areas or `for` loops. It would be more likely that one or more areas of the code accidentally getting skipped over during the modification. For this project, the time actually saved by refactoring the code is less than a second each, which is not much, but can make a significant difference when running similar code for a much larger data set.

A disadvantage of this refactored script is that if one does not understand the syntax and organization, this refactored code can be more confusing. In the process of refactoring this code, I did run into the challenge of figuring out how to best do so. This did take up a considerable amount of time to troubleshoot, so the time consumed by the process of refactoring, depending on how often the code can be applied or reused, may be an efficiency disadvantage.
