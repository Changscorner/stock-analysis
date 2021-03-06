 # Stock Analysis

## Overview of Project

The Advantages and Disadvantages of Refactoring Code.

## Purpose

1.)Taking the original VBA and editing the code to loop through all the data one time rather than once per ticker symbol

2.)Determine if refactoring increases or decreases code efficiencies.

## Dataset

[Stock Analysis](VBA_Challenge.xlsm)

## Results

### Code differences
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

The biggest difference is the code was editing the For loop to reference the arrays instead of just going through the cells. The original code looked like this:
```
    For j = 2 To RowCount
       '5a) Get total volume for current ticker
            If Cells(j, 1).Value = ticker Then

               totalVolume = totalVolume + Cells(j, 8).Value

           End If
       '5b) get starting price for current ticker
            If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then

               startingPrice = Cells(j, 6).Value

           End If
       '5c) get ending price for current ticker
            If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then

               endingPrice = Cells(j, 6).Value

           End If
    Next j
```
While the new code looks like this:
```
    For i = 2 To RowCount

        '3a) Increase volume for current ticker
            If Cells(i, 1).Value = ticker Then

                tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value

            End If

        '3b) Check if the current row is the first row with the selected tickerIndex.
            If Cells(i - 1, 1).Value <> ticker And Cells(i, 1).Value = ticker Then

                tickerStartingPrices(tickerIndex) = Cells(i, 6).Value + tickerStartingPrices(tickerIndex)
            End If

        '3c) check if the current row is the last row with the selected ticker
            If Cells(i + 1, 1).Value <> ticker And Cells(i, 1).Value = ticker Then

                'store the tickerEndingPrice into the array
                tickerEndingPrices(tickerIndex) = Cells(i, 6).Value + tickerEndingPrices(tickerIndex)

        '3d) Increase the tickerIndex.

                tickerIndex = tickerIndex + 1

            End If

    Next i
```

The new code uses the 3 output arrays that we created earlier, which allows the code to run through the worksheet only once. Under each If function in the new code the tickerVolumes, tickerStartingPrices, tickerEndingPrices each referenced their own output arrays to store the data.

### Run Time Differences
The original code has a runtime for the years 2017 and 2018 of:

<img src="https://github.com/Changscorner/stock-analysis/blob/master/Resources/Original%202017.png">
<img src="https://github.com/Changscorner/stock-analysis/blob/master/Resources/Original%202018.png">

 While the new code for the same years has a runtime of:

 <img src="https://github.com/Changscorner/stock-analysis/blob/master/Resources/VBA_Challenge_2017.png">
 <img src="https://github.com/Changscorner/stock-analysis/blob/master/Resources/VBA_Challenge_2018.png">

There is a difference of around 0.086 and 0.0625 seconds for 2017 and 2018 respectively.
This isn't even accounting for the fact that in the old code the formatting has to be run on its own afterwards instead of being packaged in.

### Stock Performance 2017 VS 2018
In the year 2017 and 2018 the returns looked like this:
<img src="https://github.com/Changscorner/stock-analysis/blob/master/Resources/Returns%202017.png">

<img src="https://github.com/Changscorner/stock-analysis/blob/master/Resources/Returns%202018.png">

As you can see in the year 2017 all of the stocks with the exception of TERP has a positive return for the year vs in 2018 where only 2 of the 12 stocks had a positive return.

## Summary

### Advantages and Disadvantages to Refactoring
#### Advantages
1.)Refactoring code is important because it allows the code to become more efficient and manageable.

2.)It makes the code more readable, run faster, and allows for an incremental approach to coding.

3.)Increases productivity

#### Disadvantages
1.) If the application's code that you are editing is too complicated it can lead to there not being proper test cases.

2.) There is a big time investment that is needed and could make the code more complex.

3.) If during the refactor additional requirements are added in it can lead to additional time to be spent on testing.

### VBA Script Refactoring

#### Advantages
The advantage for refactoring the VBA code is that it made it much more efficient with a lower execution time. The previous code ran through the worksheet 12 time for 12 outputs. While the new code ran through the worksheet only once and separated the outputting to its own for loop.

#### Disadvantages
The code became considerably more complex with the involvement of the 3 output arrays which led to some errors on the output initially. The output code had to be edited to show only the endingPrices and startingPrices only to see where the error was occurring from. Additionally it was easier to spot mistakes in the code originally as the outputs were displayed immediately rather than being stored inside an array. To debug an error in the array I would've had to
