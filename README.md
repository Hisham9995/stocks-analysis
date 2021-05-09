# stocks-analysis
green_stocks.xlsm
##Overview Project
###Purpose
The Purpose of this project was to help Steve  to find the total daily volume and yearly return for each stock.in the year 2017 and 2018   
If steve made profit or lost money  ,By Using a Microsoft Excel VBA to collect these Data , This Challenge is  to refactor the code  and run Faster.

###The Data
It represnts two charts for  2018 and 2017 these charts contains  12 different stocks with 8 columns  ticker value, the date the stock was issued, the opening,
closing and adjusted closing price, the highest and lowest price, and the volume of the stock , this challenge is to determine stock performance and the return 

## Results
### Analysis


    I did create output tickerindex = 0  and three  arrays by using Dim to declare variables ,I did use for loop to get tickerstarting Prices , tickerending Prices 
    and tickerVolumes  I also for loop helped me in counting the rows  on the other hand increasing the Volumes I used  Cells(i - 1, 1).Value <> tickers(tickerIndex)
    to determine tickerstarting Prices and Cells(i + 1, 1).Value <> tickers(tickerIndex) determining tickerending Prices in point "4" I did use for loop to get the return 
    and ticker volumes and tickers 
    
    
    
    
    '1a) Create a ticker Index
    tickerIndex = 0

    '1b) Create three output arrays
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    
    '2a) Create a for loop to initialize the tickerVolumes to zero.

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
        End If
        
        '3c) check if the current row is the last row with the selected ticker
        'If  Then
         If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
         End If

            '3d Increase the tickerIndex.
             If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
                tickerIndex = tickerIndex + 1
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

## Summary
### Pros and Cons of Refactoring Code
The Cons is debugging I had hard time figuring out errors and solving the problem another cons is sometimes the system lose fuctional 
the pros it view your project more orgnized after refactoring and clear  with a good design and look  and will be runing slightly faster 

### The Advantages of Refactoring Stock Analysis
The Advantages are to make the  code cleaner and more organized ,easy to read and analyze  , Also more simple  to fix  
### The disadvantages Refactoring Stock Analysis
It's risky when the application is big becausse there is a chance of  Functional Loss on the other hand it is time consuming and  full debugging problems 

_________{}you can find images bellow{} ___________
2017
https://github.com/Hisham9995/stocks-analysis/blob/main/2017.png

2018
https://github.com/Hisham9995/stocks-analysis/blob/main/2018.png
