# Stock Analysis Using VBA

## An Analysis of Green Energy Stock Data

### Background

Steve's parents want to diversify their investment portfolio by investing in the nascent green energy sector of the stock market. Given the volatility of the new sector, they wanted to know how a group of green energy stocks have historically performed. They've tasked Steve with finding out which ones had positive or negative returns in 2017 and 2018. Steve has asked for my help in performing a few quick analyses, so keep his parents informed on the performance of their portfolio.

### Purpose

The purpose of this analysis was to help Steve's parents determine which green energy stocks to invest in. Using ticker, daily pricing (high, low, closing prices), and daily volume from 12 different **green** stocks, I calculated the total daily volume and and annual return for each stock using a VBA script in Excel. Steve specified that he wanted to expand the dataset to include the entire stock market. To do so required a refactored VBA script, one that would provide the same functionality as my first script but run faster. This analysis will not only help Steve's parents determine which stocks invest their money but also allow Steve to use my model for quick, future analysis on thousands of stocks. The excel workbook with both VBA scripts can be found here: [VBA_Challenge](https://github.com/dwwatson1/stock-analysis/blob/main/VBA%20Challenge.xlsm)

## Results and Analysis

### Refactored VBA Script Used in this Analysis
  
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
            
        End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row's ticker doesn't match, increase the tickerIndex.
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
    
    'Formatting and Conditional Formatting
    
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

### Analysis of 2017 Green Energy Stock Data

#### Overview

To complete the analysis of 2017 green energy stock data for Steve and his parents, I need to create a VBA script to loop through the data of the 12 stocks provided. I'll refer to the first script as the __original script__ in this section. To meet Steve's expectation of being able to use my model in the future for thousands stocks, I needed to create a more efficient, refactored VBA script. I'll refer to the second script as the __refactored script__ in this section.

#### Process and Results

Using the stock data from 12 green energy stocks, I built my __original script__ run an analysis on all tickers for 2017. 

### Analysis of 2018 Green Energy Stock Data

#### Overview

#### Process and Results

## Summmary

### Advantages and Disadvantages of Refactoring Code

### Advantages and Disadvantages of Refactoring Code from this 'Analysis of Green Energy Stock Data'
