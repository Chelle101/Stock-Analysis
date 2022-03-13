# Stock Analysis with VBA + Excel
## Overview: VBA stock Analysis
## The Background
At the click of a button Steve can analyze an entire worksheet. Wanting to do a little more research for his parents, he would like to expand the dataset to include the entire stock market for the past year. Although the prior dataset was a success for dozens of stock, accounting for thousands of stock can be more challenging. However if it is accomplished the runtime is expected to be long.
## Purpose
For this analysis we are tasked with editing/refactoring the stock market dataset given in module 2 by using VBA solution code to loop through all given data one time while evaluating the whole entire stock market dataset. Which would ultimatley help us find out if refactoring our the given code was more effiecent in making the code run faster. The goal is to improve the code by making it more efficent by minimizing the steps taken to get output, reduce the amount of storage being utilized and improving the accuracy of the code by making it easy to read.
## Results
At the end of this analysis, important information such as the run time for each year needed to be extracted. The worksheets was formatted for easier visualization. In order to make the code more efficient, the nesting order of the for loop needed to be switched. In order for that to be done, four different arrays was created as listed below.
- Tickers - used to establish the ticker symbol of a stock. 
- Ticker Volumes- 
- Ticker Starting Prices
- Ticker Ending Prices

The remaining three arrays with tickers array were matched by using a variable called the tickerIndex.The steps were then listed out in order to set the structure for the refactoring code. Once complete, the analysis showed only a minor difference in the outcome as a result of refactoring the code.

- Timer 2017 before refactor .171825
- Timer 2017 after refactor .078125
- Timer 2018 before refactor .0703125
- Timer 2018 after refactor .046875
As shown above, the 2017 and 2018 timer ran faster after the changes.
### Timer before refactor ![refactor2017](https://user-images.githubusercontent.com/99842026/158042095-07568e42-5e71-45d6-ad70-888e2ab77cb7.png)
### Timer after refactor![2017 runtime](https://user-images.githubusercontent.com/99842026/158042077-554621df-001a-47a4-bef5-68d12ea3c6d2.png)
### Timer before refactor ![2018 runtime](https://user-images.githubusercontent.com/99842026/158042109-caf19323-86ca-43da-a167-aef3fbaf5a7a.png)
### Timer after refactor ![refactor2018](https://user-images.githubusercontent.com/99842026/158042121-b51d7e3c-8e2c-4169-836f-2ecd85e2a4c6.png)
# Summary
Refactoring code is essentially editing the existing code to determine if it will run faster. In order for it to properly work, it needs to be done in small steps. The steps are completed without major changes to existing code (e.g., external behavior, functionality).

## Disadvantages and advantages of refactoring
Refactoring codes can have its advantages as well as disadvantages. The following are the advantages and disadvantages that caught my attention:

The advantages of refactoring a code is that you get to add/delete and improved the affectiveness of any code. Like most programs, the more data you are dealing with, the harder the program works to analyze it all. The height of Excel's abilities with large amounts of data comes sooner than other related programs. It is important to make VBA's code as efficient as possible paying attention to every single detail in order to keep the spreadsheet working as intended. When you refactor it also helps to improve the legitmacy of the code, and helps future coders to understand what is happening with existing code.

One of the obvious disadvantages that I can think of with refactoring a code is the potential to break the existing program while working on it. However that can be easily fixed by creating backups of the data and the program, and conducting the refactoring on those backups. Another disadvantage would be the time cosumption ( It may or may not be worth it). Lastly another disadvantage would be inadvertly introducing unwanted bugs.
