# stock-analysis

## Overview of Project

### Purpose
The current code allows the user to analyze a few stocks at once but may not work as well when expanded to thousands of stocks. To test this I will refactor the exisitng code which will allow the user to analyze the entire stock market and then determine if the refactor decreased script run time.

## Results
Some of the changes I made was to run the all of the code including the formatting in the same subroutine whereas the original code had them seperated into seperate subroutines. 

Some of the code that that I made were creating a tickerIndex instead of calling specific cells as I had in the original code: 

Original Code: 
  If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker 

Refactored Code: 
"tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value" 
This allowed volume of the current "tickerVolumes" to increase by using the "tickerIndex" variable.

The results are that the refactored code may look nicer and easy to follow however it takes longer to run.

Refactored 2017 Results:
![VBA_Challenge_2017.png](/resources/VBA_Challenge_2017.png)

Original 2017 Results:
![VBA_Challenge_2017_Original.png](/resources/VBA_Challenge_2017.png)

Refactored 2018 Results:
![VBA_Challenge_2018.png](/resources/VBA_Challenge_2018.png)

Original 2018 Results: 
![VBA_Challenge_2018_Original.png](/resources/VBA_Challenge_2018_Original.png)

## Summary

- What are the advantages or disadvantages of refactoring the code?
The refactored code completes quicker than the the orginal code however, it is more difficult to write and takes more time. That being said the refactored code is nicer to look at and easier to follow.

- How do these pros and cons apply to refactoring the original VBA script?
Refactoring the code is a time investment with the benefit of the code running smoothly and quicker. The refactored code is also easier to follow and understand.
