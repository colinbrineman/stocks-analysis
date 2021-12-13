# TITLE

Module 2 Challenge: Using Excel VBA to Optimize Analysis of Stock Market Data

## AUTHOR

Colin Brineman, M.A.

## OVERVIEW

### OBJECTIVES

The challenge for Module 2 is to create a macro in Excel VBA for Steve, a recent graduate with a degree in finance, so that he can best advise his parents on how to diversify their retirement investments. During the Module 2 lessons, a macro was created for Steve which satisfied his initial goal, namely, to summarize the annual total volumes and rates of return for 12 green energy stocks. The purpose of the challenge is to refactor the original macro to run more quickly, so that Steve can put the refactored macro to use analyzing a larger data set covering the entire stock market.

### DELIVERABLES

The deliverables for the Module 2 challenge include:
1. an Excel workbook containing a refactored macro to analyze stock market data;
2. screenshots showing the runtimes of the refactored macro for the 2017 and 2018 datasets; and
3. a written report, which includes a comparison of the performances of 12 green energy stocks for 2017 vs. 2018, a description of the differences between the original and refactored macros, and an analysis of the advantages and disadvantages of refactoring code in general.

## FINDINGS

### COMPARING STOCK PERFORMANCES FOR 2017 VS. 2018

FIGURE 1: TOTAL DAILY VOLUMES AND RETURNS, 2017
![FIGURE 1: TOTAL DAILY VOLUMES AND RETURNS, 2017](/Resources/VBA_Results_2017.png)

FIGURE 2: TOTAL DAILY VOLUMES AND RETURNS, 2018
![FIGURE 1: TOTAL DAILY VOLUMES AND RETURNS, 2018](/Resources/VBA_Results_2018.png)

FIGURE 3: PERCENTAGE CHANGE IN TOTAL DAILY VOLUMES AND RETURNS, 2017 VS. 2018
![FIGURE 1: TOTAL DAILY VOLUMES AND RETURNS, 2018](/Resources/VBA_Results_Comparison.png)

Comparing Figure 1 and Figure 2, one can see that none of the stocks which had positive annual returns for 2017 went onto have positive annual returns for 2018. Further statistical analysis shows that the median return for the 12 selected stocks was 41.5% in 2017, but their median return dropped to -12.0% in 2018. From Figure 3, one can see that the stock with the smallest absolute percentage change in annual return between 2017 and 2018 had a change in return of -30.8%, while the stock with the largest absolute percentage change in annual return between 2017 and 2018 had a change in return of 1,414.9%. The picture that one comes away with from these figures is one of overwhelming volatility, although it could be said that financial analysts have other measures of stock price volatility, which could be more robust than a simple comparison of annual returns from 1 year to the next. Nevertheless, one might recommend to Steve that he advise his parents to invest in stocks with more consistent performances than these 12 green energy companies, especially as they approach retirement.

### COMPARING RUNTIMES FOR THE ORIGINAL MACRO VS. THE REFACTORED MACRO

FIGURE 4: RUNTIME FOR ORIGINAL MACRO, 2017
![FIGURE 1: TOTAL DAILY VOLUMES AND RETURNS, 2017](/Resources/VBA_Original_2017.png)

FIGURE 5: RUNTIME FOR ORIGINAL MACRO, 2018
![FIGURE 1: TOTAL DAILY VOLUMES AND RETURNS, 2017](/Resources/VBA_Original_2018.png)

FIGURE 6: RUNTIME FOR REFACTORED MACRO, 2017
![FIGURE 1: TOTAL DAILY VOLUMES AND RETURNS, 2017](/resources/VBA_Challenge_2017.png)

FIGURE 7: RUNTIME FOR REFACTORED MACRO, 2018
![FIGURE 1: TOTAL DAILY VOLUMES AND RETURNS, 2017](/Resources/VBA_Challenge_2018.png)

Comparing Figures 4 and 5 with Figures 6 and 7, one can see that there is a striking difference between the runtimes of the macros before and after refactoring. The average runtime for the original macros was 0.54296875 seconds, while the average runtime for the refactored macros was 0.09375 seconds, meaning that refactoring the macro reduced the average runtime by more than 80%. Therefore, one could certainly advise Steve to make use of the refactored macro, rather than the original one, because such a large difference in runtime could certainly make his goal of analyzing the entire stock market in Excel much quicker and easier to accomplish.

### DIFFERENCES BETWEEN THE ORIGINAL MACRO AND THE REFACTORED MACRO

#### STEPS IN REFACTORING THE MACRO
Text files of the original macro and the refactored macro can be found at [/resources/VBA_Challenge_Script.vb](/Resources/VBA_Challenge_Script.vb) and [/resources/VBA_Original_Script.vb](/Resources/VBA_Original_Script.vb) respectively. The major steps of refactoring the macro were as follows:
1. introducing a new index variable, tickerIndex;
2. changing the output variables into arrays, namely, tickerVolumes(), tickerStartingPrices(), and tickerEndingPrices(); and
3. using tickerIndex to access the correct index in the output variables and in tickers().

#### BENEFIT AND COST OF REFACTORING THE MACRO
The major benefit of refactoring the code is that the use of a common index variable means that the subroutine can find the correct values to feed to our output worksheet, for all variables simultaneously as it loops over the rows, thus meaning the loop only has to be performanced 1x, rather than 12x. The major cost of refactoring the code is that it is rather time-consuming, whereas starting from scratch could likely have been easier.

### THE ADVANTAGES AND DISADVANTAGES OF REFACTORING CODE

#### ADVANTAGES OF REFACTORING CODE

The major advantages of refactoring code are that it can:
1. help to optimize performance, thus allowing the analyst to get more work done in a given amount of time and with a given amount of processing capacity; and that it can
2. make code more flexible and adaptable, thus allowing the analyst to change the code more easily to ask different questions or to work with different datasets.

#### DISADVANTAGES OF REFACTORING CODE
The major disadvantages of refactoring code are that it can:
1. cause the analyst to make unnecessary mistakes when lines of code are copied-and-pasted, thus bogging the analyst down with fixing bugs or even introducing errors into the final results of the analysis; and that it can
2. hinder the analyst from thoroughly thinking through the relevant questions which their analysis seeks to answer, as the analyst may be find their reasoning constrained by the design of the original code.

## CONCLUSION
The Module 2 challenge, writing and refactoring an Excel VBA macro to analyze stock market data, demonstrates both how frustrating it can be to refactor code, as well as how substantial the performance gains of refactoring can be. On the one hand, the analysis performed for Steve does seem to suggest that his parents should look for other stocks to buy than the 12 green energy stocks he selected. Fortunately, however, Steve now has at his disposal a macro which is much more up to the task of analyzing the entire stock market than the original macro which was made for him during the Module lessons.
