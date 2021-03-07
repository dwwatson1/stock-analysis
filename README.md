# Stock Analysis Using VBA

## An Analysis of Green Energy Stock Data

### Background

Steve's parents want to diversify their investment portfolio by investing in the nascent green energy sector of the stock market. Given the volatility of the new sector, they wanted to know how a group of green energy stocks have historically performed. They've tasked Steve with finding out which ones had positive or negative returns in 2017 and 2018. Steve has asked for my help in performing a few quick analyses, so keep his parents informed on the performance of their portfolio.

### Purpose

The purpose of this analysis was to help Steve's parents determine which green energy stocks to invest in. Using ticker, daily pricing (high, low, closing prices), and daily volume from 12 different **green** stocks, I calculated the total daily volume and and annual return for each stock using a VBA script in Excel. Steve specified that he wanted to expand the dataset to include the entire stock market. To do so required a refactored VBA script, one that would provide the same functionality as my first script but run faster. This analysis will not only help Steve's parents determine which stocks invest their money but also allow Steve to use my model for quick, future analysis on thousands of stocks. The excel workbook with both VBA scripts can be found here: [VBA_Challenge](https://github.com/dwwatson1/stock-analysis/blob/main/VBA%20Challenge.xlsm).

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

To complete the analysis of 2017 green energy stock data for Steve and his parents, I need to create a VBA script to loop through the data of the 12 stocks provided. I'll refer to the first script as the [Original_Script](https://github.com/dwwatson1/stock-analysis/blob/main/Resources/Original_Script.bas) in this section. To meet Steve's expectation of being able to use my model in the future for thousands stocks, I needed to create a more efficient, refactored VBA script. I'll refer to the second script as the [Refactored_Script](https://github.com/dwwatson1/stock-analysis/blob/main/Resources/Refactored_Script.bas) in this section.

#### Process and Results

Using the stock data from 12 green energy stocks, I built my __original script__ to run an analysis on all tickers in 2017. As you can see from the code, I created a format in the output of the **All Stocks Analysis** sheet within [VBA_Challenge](https://github.com/dwwatson1/stock-analysis/blob/main/VBA%20Challenge.xlsm). Then, I initialized an array of all 12 tickers and variables for starting price and ending price.  After establishing loops through the data, I was able to extract data from each ticker to show **Total Daily Volume** and **Return**.  I was able to assign this macro to a button called **Year Analysis**. I was then ready to run my __original script__ script. After running it, a popup box appeared asking __What year would you like to run the analysis on?__ I specified that I wanted to see 2017 data only. 

![Year_Analysis_Button](https://github.com/dwwatson1/stock-analysis/blob/main/Resources/Year_Analysis_Button.PNG)

Because I added another spcification in my code to start and stop a timer, as well as a message box displaying the total time it took to run it, a message box popped up after I ran it.

![Year_Analysis](https://github.com/dwwatson1/stock-analysis/blob/main/Resources/Year_Analysis.PNG)



To display the output neatly, I added a script bold the column headings, add commas to the **Total Daily Volume** column, and show only one decimnal place for the percentage of **Return**.  I added a conditional formatting script to color the **Return** to show green for a positive return and red for a negative return. I assigned those two separate macros to buttons labeled **Add Coloring** and **Add Formatting**



Code in script:

MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)
 
 Message box display:
 
![2017_Stock_Macro](https://github.com/dwwatson1/stock-analysis/blob/main/Resources/2017_Stock_Macro.PNG)

The 2017 results took just under 0.8 seconds to display positive returns for every stock but $TERP:

![2017_Refactor_Analysis](https://github.com/dwwatson1/stock-analysis/blob/main/Resources/2017_Refactor_Analysis.PNG)

$DQ was the top performing stock, gaining nearly 200%. The lowest positively returning stock was $RUN, up 5.5% for the year. Overall, 2017 was an impressive year for green energy stocks. If we compare each stock's performance to a benchmark like the S&P 500, we can see if Steve's parents would just be better off investing in a less volatile mututal fund like $FXAIX. The return for the Fidelity 500 Index Fund was just under +22% for 2017. Impressively, 9 green energy stocks had better returns in 2017 than this benchmark.

I ran the same analysis again using my __refactored script__. It condensed all of the separate subs in my __original script__ into one cohesive piece of code. The full code is displayed above in the **Refactored VBA Script Used in this Analysis** section. As a result, it cut the run time by 65% to 0.27 seconds. The refactored script will come in handy for Steve when he wants to analyze thousands of stocks because he will save a lot of time waiting for the code to run. 

![VBA_Challenge_2017.PNG](https://github.com/dwwatson1/stock-analysis/blob/main/Resources/VBA_Challenge_2017.PNG)

To make it even easier for Steve, I assigned the refactored script macro to a button called **Refactor Code**. Now Steve can analyze thousands of stocks with the press of a button!

![Refactor_Button.PNG](https://github.com/dwwatson1/stock-analysis/blob/main/Resources/Refactor_Button.PNG)

### Analysis of 2018 Green Energy Stock Data

#### Overview

I repeated the same steps used in the 2017 analysis for the 2018 analysis of green energy stock data. Again, I used two scripts: the [Original_Script](https://github.com/dwwatson1/stock-analysis/blob/main/Resources/Original_Script.bas) and the [Refactored_Script](https://github.com/dwwatson1/stock-analysis/blob/main/Resources/Refactored_Script.bas) in this section.

#### Process and Results

I repeated running the two scripts but specifying the year **2018** in the popup box __What year would you like to run the analysis on?__ Similaryly to 2017 results, the __original script__ took longer than the __refactored script__. The time difference was .74 seconds compared to .26 seconds.

__Original script__ timing

https://github.com/dwwatson1/stock-analysis/blob/main/Resources/2018_Stock_Macro.PNG

__Refactored script__ timing

https://github.com/dwwatson1/stock-analysis/blob/main/Resources/VBA_Challenge_2018.PNG

The 2018 returns for the 12 green energy stocks were mostly negative, except for $ENPH and $RUN, which saw positive returns over 80%. The two also saw positive returns in 2017 of 130% and 5.5%, respectively. $TERP was the only stock that had negatives returns for both years. 

![2018_Refactor_Analysis](https://github.com/dwwatson1/stock-analysis/blob/main/Resources/2018_Refactor_Analysis.PNG)

$ENPH saw the best return from 2017 - 2018. As investor myself thinking long-term, I wouldn't necessarily feel comfortable recommending Steve's parents invest in a single stock based on two years of data. In the history of the stock market, 2 years is miniscule and putting all your eggs in one stock's basket is dangerous. $ENPH has had a great run in 2020 but after hitting a February 2021 high, the stock had sold-off significantly. 

![ENPH_Stock](https://github.com/dwwatson1/stock-analysis/blob/main/Resources/ENPH_Stock.PNG)

## Summmary

### Advantages and Disadvantages of Refactoring Code

Refactoring code is advantageous because it can take the same output and make the process more neat and efficient. By neat, I mean formatted and organized in a series of steps. This allows me (or even Steve) to look at the code and understand the process. This is helpful for a VBA novice like me. I'm still trying to wrap my head around the logic of each step, but because I'm process-oriented, the organization keeps me on track. 

The disadvantages to refactoring code is that it can be time consuming to create and at times, you may have no idea how long it'll take you to complete the process. The refactoring process could end up getting you stuck or breaking the code. Given how complex the code is, it can be overwhelming debug things, especially for someone like me who is just learning VBA.

### Advantages and Disadvantages of Refactoring Code from this 'Analysis of Green Energy Stock Data'

I learned a lot by going through the process refactoring code from my Analysis of Green Energy Stock Data. By creating an organized, step-by-step, guide within the refactored code, the logic of the individual pieces of code started to click with me. The cohesiveness  
