
# Stock Analysis Using VBA

## Overview of Project

In this project, I was tasked with analyzing green stocks to find total daily volume and annual return. I used VBA and Excel to create subroutines in order to analyze the dataset, which can also run for either given year. I also included a timer in order to time the amount of time it takes for each script to run. Then I refactored the script in order to make the script run faster. 

## Results

To make the code more efficient, I looped through all the data one time in order to collect the information. I did this by creating 3 arrays to hold tickerVolumes, tickerStartingPrices, and tickerEndingPrices. These arrays store data for each stock when the for loop runs on each. 

###### Refactored 

```
Sub AllStocksAnalysisRefactored()
    Dim startTime As Single
    Dim endTime  As Single

    yearValue = InputBox("What year would you like to run the analysis on?")

    startTime = Timer
    
    'Format the output sheet on All Stocks Analysis worksheet
    Worksheets("All Stocks Analysis").Activate
    
    Range("A1").Value = "All Stocks (" + yearValue + ")"
    
    'Create a header row
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"
    
'Initialize array of all tickers
    Dim tickers(12) As String
    
    tickers(0) = "AY"
    tickers(1) = "CSIQ"
    tickers(2) = "DQ"
    tickers(3) = "ENPH"
    tickers(4) = "FSLR"
    tickers(5) = "HASI"
    tickers(6) = "JKS"
    tickers(7) = "RUN"
    tickers(8) = "SEDG"
    tickers(9) = "SPWR"
    tickers(10) = "TERP"
    tickers(11) = "VSLR"
    
    'Activate data worksheet
    Worksheets(yearValue).Activate
    
    'Count the number of rows to loop over
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
    '1a) Create a ticker index to reference proper ticker in the arrays.
    Dim tickerIndex As Integer
    'Initiate tickerIndex at zero.
    tickerIndex = 0
    
    '1b) Create three output arrays
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    
    '2a) Create for loop to analyze each ticker in the array.
    For tickerIndex = 0 To 11
    'Initiate each ticker's volume at zero.
    tickerVolumes(tickerIndex) = 0
    
    'Activate data worksheet
    Worksheets(yearValue).Activate
        
        '2b) Loop over all the rows in the spreadsheet.
        For i = 2 To RowCount
        
            '3a) Increase volume for current ticker.
            tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
            '3b) Check if the current row is the first row with the current ticker.
                    
            If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
                'if it is the first row for current ticker, set starting price.
                tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
            
            'End If
            End If
            
        '3c) Check if the current row is the last row with the current ticker.

            If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
                'if it is the last row for current ticker, set ending price.
                tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
            
            'End if
            End If
            
        '3d) Check if the current row is the last row with the current ticker.
            If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
                
                'if it is, increase tickerIndex to move on to next ticker in array.
                tickerIndex = tickerIndex + 1
            
            'End If
            End If
    
        Next i
        
    Next tickerIndex
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    
    For i = 0 To 11
        
        'Activate Output Worksheet
        Worksheets("All Stocks Analysis").Activate
        
        'Ticker Row Label
        Cells(4 + i, 1).Value = tickers(i)
        
        'Sum of Volume
        Cells(4 + i, 2).Value = tickerVolumes(i)
        
        'ReturnValue
        Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
        
    Next i
    
'Formatting
    Worksheets("All Stocks Analysis").Activate
    Range("A3:C3").Font.FontStyle = "Bold"
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Range("B4:B15").NumberFormat = "#,##0"
    Range("C4:C15").NumberFormat = "0.0%"
    Columns("B").AutoFit

    dataRowStart = 4
    dataRowEnd = 15

    For i = dataRowStart To dataRowEnd
        
        If Cells(i, 3) > 0 Then
            
            Cells(i, 3).Interior.Color = vbGreen
            
        Else
        
            Cells(i, 3).Interior.Color = vbRed
            
        End If
        
    Next i
 
    endTime = Timer
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)

End Sub
        
```

###### Original 

```
Sub AllStocksAnalysis()

    Dim startTime As Single
    Dim endTime As Single

    yearValue = InputBox("What year would you like to run the analysis?")

        startTime = Timer
        
    '1) Format the output sheet on All Stocks Analysis worksheet
    Worksheets("All Stocks Analysis").Activate
    Range("A1").Value = "All Stocks (2018)"
    
    'Create a header row
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"
    
    '2) Initialize array of all tickers
    Dim tickers(11) As String
        tickers(0) = "AY"
        tickers(1) = "CSIQ"
        tickers(2) = "DQ"
        tickers(3) = "ENPH"
        tickers(4) = "FSLR"
        tickers(5) = "HASI"
        tickers(6) = "JKS"
        tickers(7) = "RUN"
        tickers(8) = "SEDG"
        tickers(9) = "SPWR"
        tickers(10) = "TERP"
        tickers(11) = "VSLR"
        
    '3a) Initialize variables for starting price and ending price
    Dim startingPrice As Single
    Dim endingPrice As Single
    
    '3b) Activate data worksheet
    Worksheets("2018").Activate
    '3c) Get the number of rows to loop over
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
    '4) Loop through tickers
    For i = 0 To 11
        ticker = tickers(i)
        TotalVolume = 0
   '5) loop through rows in the data
    Worksheets("2018").Activate
   For j = 2 To RowCount
       '5a) Get total volume for current ticker
        If Cells(j, 1).Value = ticker Then

          TotalVolume = TotalVolume + Cells(j, 8).Value

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
        
   '6) Output data for current ticker
   Worksheets("All Stocks Analysis").Activate
    Cells(4 + i, 1).Value = ticker
    Cells(4 + i, 2).Value = TotalVolume
    Cells(4 + i, 3).Value = endingPrice / startingPrice - 1
Next i

    endTime = Timer
    MsgBox "This code ran in " & (endTime - startTime) & "seconds for the year" & (yearValue)
   
End Sub
```

As a result of refactoring the code, it ran 0.6210937 seconds faster for 2017 and 0.5039062 seconds faster for 2018 stocks. 

Original

![2017](https://github.com/stephperillo/stock-analysis/blob/main/resources/Original_VBA_Run_Time_2017.png)

Refactored

![Refactored 2017](https://github.com/stephperillo/stock-analysis/blob/main/resources/VBA_Challenge_2017.png)

Original

![2018](https://github.com/stephperillo/stock-analysis/blob/main/resources/Original_VBA_Run_Time_2017.png)

Refactored

![Refactored 2018](https://github.com/stephperillo/stock-analysis/blob/main/resources/VBA_Challenge_2018.png)

Reading the data via an array has proven to be faster than reading in each cell individually in a nested loop.

I confirmed that the stock analysis outputs for 2017 and 2018 are the same for the refactored code as during the original code to ensure that the updated code truly runs faster.

![2017](https://github.com/stephperillo/stock-analysis/blob/main/resources/All_Stocks_2017.png)

![2018](https://github.com/stephperillo/stock-analysis/blob/main/resources/All_Stocks_2018.png)

## Summary

###### What are the advantages and disadvantages of refactoring code?

In general, refactoring code is important. The main advantage is that it can make the code more efficient; the new code can use less memory, take less time to run, and make it easier for others or even myself to understand or read in the future. 

Refactoring code can also have its disadvantages. There is always the possibility of messing up perfectly functioning code when trying to refactor the original. It is also possible to complicate the code or make it more difficult to understand for someone else. 

###### What are the advantages and disadvantages of the original and refactored VBA script in this project?

For this project, improving the script made it run faster. Doing so also organized it better and made it easier to comprehend. It also facilitates any possible future modifications by nesting the loops. If that time arises, it will be easier to adjust the code in one area of the code as opposed to manually changing the code in several different areas or `for` loops. It would be more likely that one or more areas of the code accidentally getting skipped over during the modification. For this project, the time actually saved by refactoring the code is less than a second each, which is not much, but can make a significant difference when running similar code for a much larger data set.

A disadvantage of this refactored script is that if one does not understand the syntax and organization, this refactored code can be more confusing. In the process of refactoring this code, I did run into the challenge of figuring out how to best do so. This did take up a considerable amount of time to troubleshoot, so the time consumed by the process of refactoring, depending on how often the code can be applied or reused, may be an efficiency disadvantage.
