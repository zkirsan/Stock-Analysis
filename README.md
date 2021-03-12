# Stock-Analysis with VBA
This project consists of an analysis of green energy stocks.

## Overview of Project
There is an excel file containing stock data to analyze. It is going to be using for building some automated tasks with the programming language of Visual Basic for Applications. 

### Purpose
The purpose of the project is to analyze some stocks with VBA, which means;
*to create new worksheets and subroutines for the analysis, 
*to write readable code, 
*loop over all tickers, 
*static and conditional formatting, 
*to make a run button, 
*run the analysis for any year, 
*measure code performance,
In addition to those, there is an using the refactoring method, which provides using less memory, more efficient and readable codes for future users. 

## Results
The project has three outputs. 
DQ Analysis
All Stocks Analysis
All Stocks Analysis Refactored
The DQ Analysis provides the stock information about the DAQO (Ticker: DQ). Those are the total daily volume and the return in 2018 with a run button. 
The All Stocks Analysis gives all tickers information, the total daily volume, and the return. There was created an InputBox with a run button accessing any year information. 
The All stocks Analysis Refactored makes the job of the All Stocks Analysis faster. For example, the code in the All Stocks Analysis ran in 2.476563 seconds for the year 2017. 
<p align="center"><img src="https://github.com/zkirsan/Stock-Analysis/blob/main/VBA_Normal_2017.png"></img></p>
The code in the Refactored Analysis ran in 1.273438 seconds for the same year. 
<p align="center"><img src="https://github.com/zkirsan/Stock-Analysis/blob/main/Challenge/Resources/VBA_Challenge_2017.png"></img></p>
The original script for the year 2018 gives the result in 2.421815 seconds. 
<p align="center"><img src="https://github.com/zkirsan/Stock-Analysis/blob/main/VBA_Normal_2018.png"></img></p>
In opposite to that, the refactoring script ran the code in 1.304688 seconds for 2018. 
<p align="center"><img src="https://github.com/zkirsan/Stock-Analysis/blob/main/Challenge/Resources/VBA_Challenge_2018.png"></img></p>
As a result, it showed that the refactoring script is faster than the original script. 
### The details of the original script
Sub AllStocksAnalysis()

    Dim startTime As Single
    Dim endTime As Single
    
    yearValue = InputBox("What year would you like to run the analysis on?")
    
    startTime = Timer

    Range("A1").Value = "All Stocks(" + yearValue + ")"

'1.Format the output sheet on the "All Stocks Analysis" worksheet

    
    'add three columns header
    
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"
    
'Formatting
    
    Range("A3:C3").Font.Bold = True
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Range("B4:B15").NumberFormat = "#,##0"
    Range("C4:C15").NumberFormat = "0.00%"
    Columns("B").AutoFit
    
        
'2.Initialize an array of all tickers
 
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
        
'3a.Initialize variables for the starting price and ending price
 
    Dim startingPrice As Double
    Dim endingPrice As Double
    
'3b.Activate the data worksheet

       
    Worksheets(yearValue).Activate
    
'3c.Find the number of rows to loop over

    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
'4.Loop through the tickers
        
    For i = 0 To 11
        ticker = tickers(i)
        totalVolume = 0

'5.Loop through rows in the data

    Worksheets(yearValue).Activate
    For j = 2 To RowCount

'5a.Find the total volume for the current ticker
    
    If Cells(j, 1).Value = ticker Then
    
        totalVolume = totalVolume + Cells(j, 8).Value
        
        End If

'5b.Find the starting price for the current ticker

    If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
    
        startingPrice = Cells(j, 6).Value
        
        End If
        
'5c.Find the ending price for the current ticker

    If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
    
        endingPrice = Cells(j, 6).Value
        
        End If
        
    Next j

'6.Output the data for the current ticker
    
    Worksheets("All Stocks Analysis").Activate
    
    Cells(4 + i, 1).Value = ticker
    Cells(4 + i, 2).Value = totalVolume
    Cells(4 + i, 3).Value = endingPrice / startingPrice - 1
      

    Next i
    
 'Formatting Results
 
    dataRowStart = 4
    dataRowEnd = 15
    For i = dataRowStart To dataRowEnd

        If Cells(i, 3) > 0 Then

            'Color the cell green
            Cells(i, 3).Interior.Color = vbGreen

        ElseIf Cells(i, 3) < 0 Then

            'Color the cell red
            Cells(i, 3).Interior.Color = vbRed

        Else

            'Clear the cell color
            Cells(i, 3).Interior.Color = xlNone

        End If

    Next i
    
    endTime = Timer
    MsgBox "This code ran in " & "" & (endTime - startTime) & " seconds for the year " & (yearValue)
    
    
End Sub

		


## Summary
