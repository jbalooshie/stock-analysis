# Module 2 Challenge
This repository was created as part of a 6 month Data Analystics Bootcamp administed by George Washington University. This is the repository for the Module 2 Challenge. This challenge served as an introduction to basic coding using VBA. Topics covered including writing scripts, refactoring code, and implementing macros. 

## Overview of Project
The purpose of this analysis was to look at the performance of 12 stocks in the years 2017 and 2018, so a friend can make recommendations on which ones to buy. The friend needed a quick way to see the stocks total daily volume and total return for the year. A script was written in VBA that allows them to type in the year they would like data for, and then the script generates the results and formats them. All our friend must do is run the script, they do not have to format the data or perform any calculations.  
  
By using the script we have written, our friend can quickly get the information they need. Although this code is specific to 12 stocks our friend selected, it can be modified for a different batch of stocks. Our friend tells us which stocks they would like to see, and the code can be easily changed to generate results for a new set of stocks. This way our friend can use the same dataset, and edits the code to perform analysis of additional stocks. 

## Results

### Stock Performance

For the analysis, our friend picked 12 stocks they would like to focus on. In 2017, 11 stocks provided a positive return. This reversed in 2018, when only 2 stocks produced a positive return. See the screenshots below showing the performance of the stocks. 
  
  ![2017 Performance](https://github.com/jbalooshie/stock-analysis/blob/main/2017_Performance.PNG)
  
  ![2018 Performance](https://github.com/jbalooshie/stock-analysis/blob/main/2018_Performance.PNG)
  
Based on these results, we would advise our friend to invest in the tickers ENPH and RUN, both of which had positive returns in 2017 and 2018. However, the poor returns in 2018 may be explained by other factors, such as new challenges for alternative energy companies. There might be an opportunity for additional research into why most of the stocks performed poorly in 2018. But based on the data ENPH and RUN  produced positive returns in both years. 
  
  ### Execution Times
  
The refactored code performed significantly faster than the original script we ran. The original script took about .6 seconds to complete, while the refactored code completed the same task in .13 seconds. 

![Original Code](https://github.com/jbalooshie/stock-analysis/blob/main/Original_2017.PNG)

![Refactored Code](https://github.com/jbalooshie/stock-analysis/blob/main/VBA_Challenge_2017.PNG)

The main reason for the faster performance was changing how the script looped through the data. In the original code, the script looped through every row for one ticker, pulling the starting price, ending price, and total volume. It then filled the data into our output table and looped back through every row again for the next ticker. In the refactored code, it looped through every row only one time, writing the appropriate values to an array. Once it had looped through the entire sheet, it fills the data from the arrays into the output table. 

Here are two examples of how I retrieved the ending price for each stock. The first is from the original code, the second is from the refactored code. The refactored code uses a `tickerIndex` variable to write the value to the array. Once all data for that stock has been gathered, the `tickerIndex` variable increases by one, and the script continues looping through the rows, now looking for the next ticker value. 

```

   If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then

              endingPrice = Cells(j, 6).Value
```      

```
If Cells(i + 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then

               tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
             
```

# Summary

## What are the advantages or disadvantages of refactoring code?
The main advantage of refactoring code is allowing it to run faster and more efficiently. The half a second difference here might seem miniscule, but for more complex code it can be massive. The amount of data you are working with can magnify inefficiencies, like how in our original code it was running through every row 12 times. With a larger dataset this could make running our code very time consuming. Refactoring also helps sharpen your skills as a coder. Being able to look critically at your code and identify ways to make it run better ensures you can stay flexible and open-minded. 

The disadvantage to refactoring is that it can be a laborious task as well. Depending on how much code you must go through, it can take a long time to review everything, identify ways to improve it, implement the changes, and test them. If you are close to a deadline, and the code is working, it might not be worth the time to go through and review everything. Refactoring can also be difficult if you are not aware of different ways to approach the problem you are trying to solve. If there is a more elegant solution to the problem out there, but you have not seen it before, you might not always have that moment of inspiration to realize. Refactoring does seem like a good exercise to perform with a peer, reviewing each other's code can get around this disadvantage. 


## How do these pros and cons apply to refactoring the original VBA script?

For the original VBA script, we were only looking at 12 stocks. So, the difference between .6 seconds and .1 seconds is not noticeable. But if he wanted us looking at 120 stocks, then it might cause problems. Assuming increasing the number of stocks 10 times causes the code to take 10 times longer to run, it would take 6 seconds for the original script to run. The refactored code would only take one second. If you think about a phone app, anything that takes longer than a few seconds to refresh feels like an eternity. The original code we wrote would not be able to scale efficiently to look at large batches of stocks. 

The main disadvantage is that the process can be time consuming, and this impacted me. Using the `tickerIndex` variable was  a creative way to get the script to run faster. It took me several hours before I got it running correctly. My main issue was confusing locations where I needed to reference the `i` variable in the loop I was running versus locations where I needed it to reference the current `tickerIndex`. Googling, trial and error, and some help from the Thursday class session this week helped me get everything running. Today I started fresh (my old code was a mess) and was able to get it running correctly in under an hour.

