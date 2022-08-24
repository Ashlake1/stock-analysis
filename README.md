# stock-analysis

## Overview of Project

### Purpose
The current code allows the user to analyze a few stocks at once but may not work as well when expanded to thousands of stocks. To test this I will refactor the exisitng code which will allow the user to analyze the entire stock data set and then determine if the refactor decreased script run time.

## Results
One of the changes I made was to run all of the code including the formatting in the same subroutine whereas the original code had them seperated into seperate subroutines. 

A big change that was made in the refactored code was to loop the entire sheet instead of specified cells as I had done in the orignal code: 

###Original Code: 
  If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker 

###Refactored Code: 
"tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value" 
This allowed volume of the current "tickerVolumes" to increase by using the "tickerIndex" variable.

The results are that the refactored code may look nicer and easy to follow however it took a bit longer to write. That being said the, the refactored code runs quicker despite looping through a much larger data set.



## Summary

- What are the advantages or disadvantages of refactoring the code?

The refactored code took longer to write and although it takes longer to run it does go through the entire Data Set where as the orignal code only takes into account designated cells. Another benefit is that the refactored code completes quicker.

- How do these pros and cons apply to refactoring the original VBA script?

Refactoring the code is a time investment but is less taxing to run. Mostly this script is a time investment.
