# Module 2 Challenge: Stock Analysis using VBA
## Overview of the Project

### In this project, I helped Steve analyze datasets of the entire stock market over the last few years. I refactored, or edited, the Module 2 solution code to loop through all of the data one time to collect that same information that I did in the the module. The goal was to see if it is more efficient to refactor the code to do an analysis of thousands of stocks, instead of using the original code. 

## Results

By using VBA and the starter code provided in this Challenge, I refactored the Module2_VBA_Script in order to loop through the data one time and collect all of the information. The refactored code ran faster than it did in this module. I had to create three output arrays: tickerVolumes, tickerStartingPrices, and tickerEndingPrices. The script then had to loop through stock data, reading and storing all of the following values from each row. 

As you can see below, the stocks performed much better in 2017 than in 2018. One (TERP in red) out of twelve stocks did poorly in 2017. In 2018, only two (ENPH and RUN) had a positive return (green cells).
![VBA_Challenge_2017.png](https://github.com/c-ramos/stock-analysis/blob/78eb1e2f9de7e8266e55b3db760c1334f3cf4440/VBA_Challenge_2017.png)

![VBA_Challenge_2018.png](https://github.com/c-ramos/stock-analysis/blob/78eb1e2f9de7e8266e55b3db760c1334f3cf4440/VBA_Challenge_2018.png)

Additionally, the refactored code did run faster and more efficiently than the previous code. As you can see below, the original code for 2017 took much longer than it did in the refactored version. 2017 factored code took 0.2421875 seconds and the original code took 1.46875 seconds.The 2018 factored code demonstrate similar results. 

![2017_module 2_time.png](https://github.com/c-ramos/stock-analysis/blob/d98805d5446de71b51b2b90eb306a3ec43b6fbda/2017_module%202_time.png)
![2018_module 2_time.png](https://github.com/c-ramos/stock-analysis/blob/d98805d5446de71b51b2b90eb306a3ec43b6fbda/2018_module%202_time.png)
## Summary

Advantages: Refactoring cleaned up the original code and made it easier to parse through. The most significant beneift is that it takes a shorter amount of time to run the written code.

Disadvantages: Refactoring takes a bit more time. I can see this being a problem if the dataset is very large which can take more time. In this case, since the Module 2 Solution Code wasn't too large it was not too time-consuming.
