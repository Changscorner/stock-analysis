# Stock Analysis

## Overview of Project
 
Examine the reason for refactoring VBA code and the efficiencies or inefficiencies it can add.
  
## Purpose

1.)Taking the original VBA and edit the code to loop through all the data once rather than once per ticker symbol 
	
2.)Determine if this method will decrease or increase the code runtime

## Results

The original code was written to ask the user to input a year for which they would like to have analyzed using this code: 
```		
    yearValue = InputBox("What year would you like to run the analysis on?")		
```
Taking that input it would then run through the selected years worksheet once for each ticker and output the ticker, volume and returns data.
This is rather inefficient as the code has to runthorugh the worksheet 12 times so instead of doing it this way a better method would be to create 3 arrays that would store each output type for each ticker symbol.
I created 3 output arrays like this:
```
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
```
This method allows each tickers data to be stored inside its own array that can be outputted later after all of the data has been stored.
The original code has a runtime for the years 2017 and 2018 of:

<img src="https://github.com/Changscorner/stock-analysis/blob/master/Resources/Original%202017.png">
<img src="https://github.com/Changscorner/stock-analysis/blob/master/Resources/Original%202018.png">
 
 While the new code for the same years has a runtime of:

 <img src="https://github.com/Changscorner/stock-analysis/blob/master/Resources/VBA_Challenge_2017.png">
 <img src="https://github.com/Changscorner/stock-analysis/blob/master/Resources/VBA_Challenge_2018.png">
 
There is a difference of around 0.086 seconds and 0.0625 for 2017 and 2018 respectively. 
This isn't even accounting for the fact that in the old code the formatting has to be run on its own afterwards instead of being packaged in.

	
## Summary

### Advantages and Disadvantages to Refactoring
#### Advantages
1.)Refactoring code is important because it allows the code to become more efficient and managable. 

2.)It makes the code more readable, run faster, and allows for an incremental approach to coding.

#### Disadvantages
1.) There is a risk if the application's code that you are is too large or when the existing code doesn't have proper test cases.

2.) There is a big time investment that is needed and could make the code more complex.

3.) If during the refactor additional requirements are added in it can lead to additional time to be spent on testing.

There is a detailed statement on the advantages and disadvantages of refactoring code in general (3 pt).

### VBA Script Refactoring

#### Advantages

#### Disadvantages
