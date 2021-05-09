# Stock Analysis with Excel VBA
# Overview of Project
## Purpose
The purpose of this project is to refactor a Microsoft Excel VBA code and evaluate if our edits successfully make the VBA script run faster and create efficiencies. This project looked for us to create new functionality and by making the code more efficient – created fewer steps, used less memory and improved the logic of the code.
## Data
The data utilized in this project was information on 12 different stocks in 2017 and 2018. The excel file contains information on each stock’s ticker value, date of issue, the opening, closing and adjusted closing price and the volume of the stock. The goal of this data was to be able to look at the results of the 12 stocks annually to evaluate the total daily volume of stocks and the percentage return. By doing this the data would evaluate whether or not one would suggest purchasing one of the 12 stocks.
# Results
## Results of the Refactored VBA Code
By refactoring the VBA code I was able to change the original code which was good at analyzing a dozen stocks more slowly, to being able to analyze thousands of stocks in seconds to illustrate the best stock options to invest in. The new code decreased the time from nearly 0.50 seconds to run to just under 0.15 seconds. Below are screenshots of the new time it took the code the run. Additionally, from this code it showed that for 2017 data, majority of stocks had a postitive percentage return. Specifically, DQ and SEDG stocks had the highest percentage. While in 2018, all but two stocks had a negative perceentage return. Stocks RUN and ENPH had positive percentages of return.

![VBA_Challenge_2017](Resources/VBA_Challenge_2017.png)
![VBA_Challenge_2017](Resources/VBA_Challenge_2018.png)

## VBA Code
Final Refactored VBA Code

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
    
    'Get the number of rows to loop over
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
    '1a) Create a ticker Index
      tickerIndex = 0

    '1b) Create three output arrays
    
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    For i = 0 To 11
        tickerVolumes(i) = 0
        tickerStartingPrices(i) = 0
        tickerEndingPrices(i) = 0
    
    Next i
        
    ''2b) Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
            
          If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
         
            
        'End If
        End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
        'If  Then
        
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
         tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
         
        End If


            '3d Increase the tickerIndex.
            If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
                tickerIndex = tickerIndex + 1
            
        'End If
        End If
        
    Next i
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
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


# Summary
## Advantages Refactoring Code
The biggest advantage to refactoring code is that is creates a more organized, concise with a simpler set up. The updated code can be useful for debugging, software and design improvements and faster speed. Because it is a more efficient set up it is easier to follow and more straightforward.
## Disadvantaages Refactoring Code
A large disadvantage of refactoring code is that it is not always an option. This is due to some data sets not being easily changeable with other refactored code. It is possible that while you are refactoring code you make an error that changes the outcome of the code and alters the results. 
## Advantages of Original and Refactored VBA script
In this case, the biggest advantage for refactoring this VBA script was decreasing the run time. Initially when we first ran the code both 2017 and 2018 data took just under 0.50 seconds. However, when using the refactored code 2017 and 2018 data was able to be run in under .16 seconds, screenshots below. 
## Disadvantages of Original and Refactored VBA script
A slight disadvantage of refactoring the VBA script was the time it took to fully understand each aspect of the old code in order to ensure the updated code was looking to analyze the same details and results. 

