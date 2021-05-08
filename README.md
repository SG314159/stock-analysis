# stock-analysis
Stock analysis - Challenge 2 for UT Bootcamp


## Overview of Project
The purpose of this project is to analyze stock data from green energy companies and make a recommendation for future investment. The data consists of two Excel workbooks with data on 12 stocks for the years 2017 and 2018. The analysis was performed via VBA scripts in Excel 2017, running on a Windows machine. The running time for the original VBA code is compared with running time for the optimized VBA code.

## Results
### Stock Performance Comparison
According to the data, the year 2017 was a great year to have green energy stocks. Of the 12 stocks listed, only TERP had a negative return in 2017. The company DQ had an amazing 200% return. The year 2018 was the opposite situation. Only ENPH and RUN had positive returns (of about 80%), and the other 10 stocks had negative returns. If you had to select based only on the available data, choosing company ENPH is the best option. ENPH had a return of 129.5% in 2017 and a return of 81.9% in 2018.

As a reminder, past performance in the stock market is not a guarantee of future performance. It is always a possibility that investments can lose value.  


### Execution Times
When the original script ran, both the year 2017 and 2018 each took about 0.85 seconds. The refactored script ran at about 0.27 for each year. Thus, it was about half a second faster with the improved script. 


## Summary

### Question: What are the advantages or disadvantages of refactoring code?
### Question: How do these pros and cons apply to refactoring the original VBA script?
##### Advantages
The refactored code runs faster. For this data set, the difference was not major as there were only about 3000 rows for each year. For larger data sets, the faster code would produce a more noticeable  difference in performance. It is also a little clearer to follow the logic of the refactored code. Each row is examined as it is read: adding to the cumulative total as well as determining whether the row was the first or the last of that ticker symbol. In the original code, there were multiple passes through the entire data set, which made the execution time unnecessarily longer.

##### Disadvantages
The refactored code requires that the data in each worksheet already be sorted by Ticker symbol and by date. If the data were unsorted, the code would not work. Code for adding sorting could be added but would add to execution time as sorting is a computation-intensive operation. 
The refactored code also only allows for the 12 stocks included in the array definition. If another stock were considered, a programmer would have to modify the code to include the new symbols. In this sense, the code is not very robust as it is not as useful as it could be in a variety of input situations.
Both of these disadvantages occurred in the original VBA code. 
