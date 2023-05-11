# All-Stocks-Analysis
VBA Challenge Module 2
## Overview of Project
Steve is a recent finance graduate that received his first job from his parents who are passionate about green energy and want to invest in it. After reviewing his parent's initial investment, in DAQO New Energy Corporation, Steve wants to diversivy his parents stock portfolio. So, Steve asked for our assistance to analyze the stock data using Visual Basic for Applications (VBA) in Excel 

Overall, the purpose of the analysis assignment is to refactor VBA code to help Steve review data of the the green energy stock market for 2017 and 2018 more efficently. 
## Results
### Original/Starter Images
![GreenStocks 2017](https://github.com/Jall3n/All-Stocks-Analysis/assets/119149740/8b866b24-30be-435d-8791-e0b8f3fa63df) ![GreenStocks 2017 Seconds](https://github.com/Jall3n/All-Stocks-Analysis/assets/119149740/2daa5678-5383-40b1-a2cd-4031d371ff84)


![GreenStocks 2018](https://github.com/Jall3n/All-Stocks-Analysis/assets/119149740/f4184da0-b0a6-40ac-8f62-395ae29068c8) ![GreenStocks 2018 Seconds](https://github.com/Jall3n/All-Stocks-Analysis/assets/119149740/8e6fcf29-8641-4e9b-bc80-5398782ceb6e)

The original/starter code ran for 2017 in 0.2470093 seconds and ran for 2018 in 0.2440186 seconds. Additionally, 2017 had 11 positive returns but in 2018 we saw only 2 positive returns. 
One of the differences between the code and the Module 2 activity was activiating both 2017 and 2018 worksheets instead of just 2018. The solution for this came from the challenges starter code, Worksheets(yearValue).Activate, which activates and loops through the two worksheets. A similar update occurred for Range("A1").Value = "All Stocks (2018)" by changing it to "All Stocks (" + yearValue + ")". These two updates would be brought over to the refactored code as well. 

### Original/Starter Code

      Sub AllStocksAnalysis()

      '1) Format the output sheet on All Stocks Analysis worksheet
      Worksheets("All Stocks Analysis").Activate

      Dim startTime As Single
      Dim endTime  As Single

      yearValue = InputBox("What year would you like to run the analysis on?")
      startTime = Time
      Range("A1").Value = "All Stocks (" + yearValue + ")"
   
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
      '3b) Activate data worksheet; need to change to yearValue
       Worksheets(yearValue).Activate
      '3c) Get the number of rows to loop over
      RowCount = Cells(Rows.Count, "A").End(xlUp).Row

      '4) Loop through tickers
      For i = 0 To 11
          ticker = tickers(i)
          totalVolume = 0
          '5) loop through rows in the date; need to change to yearValue
          Worksheets(yearValue).Activate
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
       '6) Output data for current ticker
       Worksheets("All Stocks Analysis").Activate
       Cells(4 + i, 1).Value = ticker
       Cells(4 + i, 2).Value = totalVolume
       Cells(4 + i, 3).Value = endingPrice / startingPrice - 1
      Next i
      'Need a message box to give us our time for comparison
          endTime = Timer
          MsgBox "This code ran in" & (endTime - startTime) & "seconds for the year" & (yearValue)
      End Sub

### Refactored Code

For the challenge we needed to update the code to run more efficiently. I also combined all the formating and message box into one code. 

    Sub useAllStocksAnalysisRefactored()
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
    
1a) Create a ticker Index, set as zero

    Dim tickerIndex As Integer
    tickerIndex = 0
    
1b) Create three output arrays; for volume used Long for larger values/bits

    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    
2a) Create a for loop to initialize the tickerVolumes to zero.

    For i = 0 To 11
    tickerVolumes(i) = 0
    Next i

2b) Loop over all the rows in the spreadsheet.

    For i = 2 To RowCount
        
3a) Increase volume for current ticker

    If Cells(i, 1).Value = tickers(tickerIndex) Then
    tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
    End If
        
3b) Check if the current row is the first row with the selected tickerIndex. Utilizing <> 'doesn't equals to' and And.

    If Cells(i - 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
    tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
    End If
        
3c) Check if the current row is the last row with the selected ticker.

    'If the next row’s ticker doesn’t match, increase the tickerIndex.
    If Cells(i + 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
    tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
        
3d) Increasing the tickerIndex.

    tickerIndex = tickerIndex + 1
    End If
        
    Next i
      
4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.

       For i = 0 To 11
        Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
        
        Next i

        Formatting; include in first sub
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
### Refactored Image
The results of the refactored code show that it is faster than the original/starter code which was 0.07421875 seconds for 2017 and 0.0703125 seconds for 2018. 

<img width="265" alt="VBA Refactored 2017 Seconds" src="https://github.com/Jall3n/All-Stocks-Analysis/assets/119149740/eb0984f3-b041-430c-b760-b356912753bc">  <img width="260" alt="VBA Refactored 2018 Seconds" src="https://github.com/Jall3n/All-Stocks-Analysis/assets/119149740/1c696f60-2bb3-4cab-a9dd-a6c666317dd3">

## Summary
### Advantages and Disadvantages of Refactoring Code in General
Some advantages of refactoring code in general is the possibility of making the code more efficient and run faster.
A disadvantage of refactoring code in general is the larger amount of time it can take. When looking online it seems to be one of the top disadvantages brought up due to size of the code and the time spent testing it. 
### Advantages and Disadvantages of Original and Refactored VBA script
An advantage is that our code ran faster by around 0.17 seconds for both years while a disadvantage is how long it took to get the code to work without running into errors. Most of the time consumption came from spelling and syntax errors as well as properly setting up the new variable called "tickerIndex" and the three output array. 

1) https://www.techtarget.com/searchapparchitecture/definition/refactoring#:~:text=Refactoring%20improves%20code%20readability%20and,to%20a%20program's%20source%20code.
2) https://en.wikipedia.org/wiki/Code_refactoring
3) https://www.liquidweb.com/blog/code-refactoring/#h-what-are-the-benefits-of-code-refactoring
4) https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/data-type-summary

