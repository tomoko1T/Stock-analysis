# Searching for the right investment in the stock market

## Overview of Project 
Steve, who works for the financial industry, wants to support his parents' financial decision.  His parents are leaning to invest on the stock of only an energy company, Daqo. However Steve believes that putting their all eggs in one basket is risky decision.  He wants to advise them to diversify their portfolio by showing the solid information.  Based on his request, some stocks in the market were selected, and the analysis was run with the dataset of the total daily volume and yearly return for each stock.  The original analysis was conducted pulling the dataset from only 2017 and 2018.  However, Steve wants to take the further steps collecting the data of the entire stock market over the last few years.  We predict it could take a longer time to run the high volume of data so refactoring was processed to make VBA script run faster and efficiently.            

## Analysis Results 
In this refactoring, a ticker Index was set to zero at the beginning before iterating over all the rows which helps to cross reference the correct index while this setting was missing in the original script.  

Code : tickerIndex = 0

Data type, Single was chosen for tickerStartingPrices and tickerEndingPrices under three output arrays.

Dim tickerVolumes(12) As Long

Dim tickerStartingPrices(12) As Single

Dim tickerEndingPrices(12) As Single 

By chooing Single, it aimed to save memory consumption and support to run this script faster than the original. 

 -   VBA script is [GitHub Pages](https://github.com/tomoko1T/Stock-analysis/blob/main/VBA_Challenge.xlsm)

## Summary 
The advantage of refactoring code in general is making the program run efficiently.  The code could often be complex and difficult to be understood due to its complexity, so refactoring simplify the code and make the program run faster.  The disadvantage is although the goal of refactoring is not alternating the existing code, it may cause new bugs or errors by renaming variables and methods or moving code around.  The end result may not be positive while it takes extra time and effort to work on refactoring ; the program run slower than the original.

The advantage and disadvantages of the original and refactored VBA script are same as refactoring the code in general. In this analysis,  refactoring did help to have the original VBA script run faster.  Mainly by setting ticker Index to zero and selecting the data type, Single instead of Double.  
The image on the left side is the elapsed time of the original script, and the image on the right side is the one refactored.
Although there is the slight improvement in execution time, refactoring supported in this case analysis.  

![This is an image](https://github.com/tomoko1T/Stock-analysis/blob/main/Resouces/VBA_Challenge_2018_Original.png)
![This is an image](https://github.com/tomoko1T/Stock-analysis/blob/main/Resouces/VBA_Challenge_2018.png)
