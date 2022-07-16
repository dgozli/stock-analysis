# Analysis of Stock Data with VBA

## Overview of Project

### Purpose

This project intended to aimed to provide a performance summary of 12 companies in the stock market in 2017 and 2018. We looked at two metrics for each of the stocks, (1) total daily volume of trade and (2) the annual return. The first metric was used to gauge the overall activity of the stock, "the higher the better," and the second metric was used to gauge the overall growth or decline of the stocks over each of those two years. We were most interested in DQ, because our client's parents were interested in investing in that company. Therefore, we payed special attention to DQ and comparing it with the other 11 stocks. 

## Results

Analyses were conducted in Excel with the help of VBA scripts. We are here interested in the performance of stocks (overall, as well as DQ specifically). We are also interested in the performance of two versions of the VBA script: the original vs. the refactored code. 

### Analysis of Returns

The result showed that there was an overall growth in stocks in 2017, whereas there was an overall decline in 2018. The direction of DQ stocks in both years agreed with the majority of stocks. In 2017, the DQ had the impressive return of 199.4%. That rate of growth was, however, not an outlier and belonged to the fourth (best) quartile in our sample.

In 2018, the DQ stock also agreed with the overall pattern of decline, with -62.6% return. Again, DQ was not an outlier, though its performance belonged to the first (worst) quartile in our sample. 

### Analysis of script efficiency

We were also interested in the performance of our VBA scripts. With the original code, running the script for the years 2017 and 2018 took, respectively, 0.92 and 0.87 seconds. By contrast, the refactored script was more efficient. For both years 2017 and 2018, the refactored script took about 0.16 seconds.

## Summary

### Advantages of Refactoring Code

Refactoring code is greatly beneficial to inexperienced programmers, because it allows us to look at the code from different perspectives. Refactoring a code is like trying to solve a problem using a new method, which enriches how we understand the problem. It can also increase the elegance and economy of the code, and a more elegant code is easier to understand by others (and by ourselves in the future). Finally, in the case of our code, refactoring saved computer time, and this would be a big factor at larger scale (when we apply this code to larger datasets). 

Refactoring a script will most likely involve trade-offs. When we shorten the length of a script, or when we reduce the number of loops, it is worth asking if our new method involves its own cost. Are we defining more variables, particularly array-variables, that might take up more memory space? It is also useful to keep in mind if our code is flexible (able to be refactored later on) as we revise it. 


### Application to Refactoring the Original VBA Script

The original code involved two embedded loops, requiring VBA to go through the entire dataset 12 times (the number of tickers). Although for each of those 12 times, our VBA script was only interested in obtaining data for one of the stocks (only one "ticker"), we asked the program to check all of the rows. We changed that in the refactored code. 

In the refactored code we asked the program to monitor tickers in a continuous way, rather than fixing on a particular ticker/stock each time, which is why the code only has to go through the dataset once. This would make a big difference with larger datasets. 

Were there any disadvantages? Perhaps. In refactoring the script, we defined several array-variables, which required more short-term computer memory. In the original code, we did not use such arrays, so the demand of the original script was relatively smaller on the short-term memory. This is something we might have to consider if the arrays (or the number of array-variables) increase. 