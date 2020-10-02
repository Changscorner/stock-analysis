# Stock Analysis

## Overview of Project
 
Examine the reason for refactoring VBA code and the efficiencies or inefficiencies it can add.
  
## Purpose

1.)Taking the original VBA and edit the code to loop through all the data once rather than once per ticker symbol 
	
2.)Determine if this method will decrease or increase the code runtime

## Results

The original code was written to ask the user to input a year for which they would like to have analyzed. 
```		
yearValue = InputBox("What year would you like to run the analysis on?")		
```
Taking that input it would then run through the selected years worksheet once for each ticker signigicantly increasing the runtime to output the data.
	
	
## Summary


The purpose and background are well defined (2 pt).
Results
The analysis is well described with screenshots and code (4 pt).
Summary
There is a detailed statement on the advantages and disadvantages of refactoring code in general (3 pt).
	cons
		easier to find mistakes in the original code than in refacotred
		
There is a detailed statement on the advantages and disadvantages of the original and refactored VBA script (3 pt).
	cons
		easier to find mistakes in the orignal code than looking for it in the array
		issue i ran into was that my returns in the refactored code wasn't outputting the right number and i ended up having to test the code by having it only output the starting price and ending price alone to see if i could spot where the error was
