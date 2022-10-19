# Stock Analysis

Click here to view the Excel File: [VBA_Challenge.xlsm](https://github.com/pfrivas/Stock-analysis/blob/f515dfda59066d5db5f37cf1ad8fdaf2a1f123af/VBA_Challenge.xlsm)

---

## Overview of Project

### Purpose

The purpose of this project is to use Microsoft Excel VBA (Visual Basic for Applications) Script, which interacts with Excel using complex logic and reading cells, in order to collect specific data for stocks in the years 2017 and 2018 and in order to determine whether the specific stock (DQ) was worth investing in. The code was completed (greenstocks) in order to automate analyses that allows for reusage with different data and reduces the chances of accidents and errors. The code was then refactored (VBA Challenge) in order to improve on the efficiency of the original code.

### Background
* Steve just graduated with a finance degree and his parents are his first clients, They are currently investing in green energy stocks, specifically the DAQO's stock. Steve is looking into DAQO's stock for his parents and is concerned about diversifying their funds and wants to analyze a handful of green energy stocks in addition to DAQO stock so he created a spreadsheet containing the stock data used for this analysis. Steve's parent want to know how actively SQ was traded in 2018. They believe that if a stock is traded often, then the price will accurately reflect the value of the stock. If the daily volume for DQ is summed up then the yearly volume can be calculated and a rough idea is formed of how oftenthe stock gets traded

* Steve wants to know how well DQ performed in 2018, one way to measure that is calculate the yearly return for DQ, the yearly return is the percentage increase or decrease in price from the beginning to the end of the year. Usage of loops and conditions and checking the conditions using logical (Boolean) operators are used in this project in order to create an efficient data analysis

---

## Results

* Overall the results showed that 2017 was a better year than 2018 for stocks. The stock that Steve's parents were investing in had a positive return in 2017  (199.4%) negative return (62.8%) in 2018

* Refactoring the code in VBA allowed for the amount of time the code ran to be reduced by half a second (0.6 -> 0.1) 

### Analysis 

* Based on the results for the return and total daily volume the best course of action for Steve's parent's would be to invest in the ENPH and RUN stocks instead of the DAQO stock since those were the only 2 green energy stocks that seemed to have a positive return both in 2017 and in 2018

### Screenshots

Below is the screenshot for the 2017 stock

* <img width="672" alt="VBA_Challenge_2017" src="https://user-images.githubusercontent.com/110814780/196634036-5480b392-5fec-43ab-9c5c-a87e305de55a.png">

Below is the screenshot for the 2018 stock

* <img width="665" alt="VBA_Challenge_2018" src="https://user-images.githubusercontent.com/110814780/196634061-a9a8db7e-566b-4382-99c6-b85843c2a406.png">

### Code

* The refactored code for the tickerIndex is listed below 

    '1a) Create a ticker Index

            tickerIndex = 0

    '1b) Create three output arrays
        
        Dim tickerVolumes(12) As Long
        
        Dim tickerStartingPrices(12) As Single
        
        Dim tickerEndingPrices(12) As Single

  ''2a) Create a for loop to initialize the tickerVolumes to zero.
    
      ' If the next row’s ticker doesn’t match, increase the tickerIndex.
    
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
        
        'Output data for current ticker
        Worksheets("All Stocks Analysis").Activate
        
            Cells(4 + i, 1).Value = tickers(i)
            Cells(4 + i, 2).Value = tickerVolumes(i)
            Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
       Next i
   
---

## Summary

### Advantages and Disadvantages of Refactoring Code in General
* Advantages of refactoring code is that it is cleaner and efficient. Due to less lines of code, the code is easier to read, easier to detect bugs in, and runs the program faster time wise. Thus refactoring the code is more efficient and easier to maintain.

* Disadvantages of refactoring the code is that refactoring can become tricky when the code is too large because incorrect refactoring could lead to more errors and bugs in the code

### Advantages and Disadvantages of the Original and Refactored VBA Script
* Advantages of refactored VBA script is that there was a x5 decrease in macro run time. The original analysis took 0.6 seconds to run whereas the new analysis took about 0.1 seconds to run.

* Disadvantages of the original code is that not only is the code slower in running, but because there is more lines of code, more bugs and errors cold develop

