# stock-analysis
module-2 stock analysis VBA project


### Overview of Project: Explain the purpose of this analysis.

The purpose of this analysis was to analyze various stocks from 2017 and 2018, consolidate volumes and find yearly returns of each stock and show how each stock has performed in the current year (2018) as compared to previous year (2017) in terms of volumes and returns. 
### Results: Using images and examples of your code, compare the stock performance between 2017 and 2018, as well as the execution times of the original script and the refactored script.

#### stock performance 2017 vs 2018

The stocks in 2017 did really well as compared to 2018. There were heavy returns, some in triple digits, when it came to performance. There was only a single stock TERP that had a negative return. You can see the details in the screenshot below.  

![ScreenShot](https://github.com/LIPSASHARMA/stock-analysis/blob/9a445f3316d145cd907e4d7c20879756df73bdaa/Resources/VBA_Challenge_2017.png)

The 2018 is kind of opposite to what we have seen in 2017. All the stocks gave a negative return except ENPH and RUN stocks, which gave significant returns 0f 81% and 84% respectively. You can see the details in the screenshot below.

![ScreenShot](https://github.com/LIPSASHARMA/stock-analysis/blob/9a445f3316d145cd907e4d7c20879756df73bdaa/Resources/VBA_Challenge_2018.png)


#### Execution times between original script and refactored scripts

The initial execution time for 2017 and 2018 was 0.9492188 and 0.9375 respectively, which is quite significant looking at the number of records. If there is a need to analyze millions of records then this processing times will significantly increase. 

On the other hand, the refactored script reduced the processing times by 8-9 times. The refactored code resulted in 0.171875 seconds and 0.1640625 seconds for 2017 and 2018 respectively. This is a significant improvement over the previous code. 

The primary reason this happened because we managed to use a single FOR-loop as compared to two nested loops earlier and also use arrays for tickerVolumes, tickerStartingPrices and tickerEndingPrices. This helped in reducing the execution time of the code. 


### Summary: 

#### 1.	What are the advantages or disadvantages of refactoring code: 

The biggest advantage of having a refactored code is that the execution time reduces drastically thereby increasing its performance, however, some times it becomes difficult to understand a highly refactored code unless it is backed up with a detailed documentation and comments

#### 2.	How do these pros and cons apply to refactoring the original VBA script: 

The execution time has reduced from 0.8 seconds to 0.1 seconds but now it is become a bit tricky to understand the function/working of tickerIndex variable (in my personal opinion).


