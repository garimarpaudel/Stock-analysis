# Green Stock analysis.
## 1. Overview
#### Steve's parents wants to invest in the Green Energy Stocks,however,Steve thinks its not a good idea to invest all the money in just one stock without analyzing the stock market data for Green Stocks performance in prior years.Therefore, Steve needs our help to create a VBA program that automates the stock data analysis for past, present and future. 

## 2. Results
#### Two different methods were adopted to compare the performance of the stock analysis bases on VBA code run and performance time generating the same result for stock analysis. The original VBA code recommended in Module 2 and refactored VBA code were compared for their perfomances in analyzing data. 

#### 2.1 Module 2 recommended VBA code 
#### The result of stock data analysis for year 2018 using module 2 recommended VBA code is presented below:
![Module 2 result 2018](https://user-images.githubusercontent.com/108683284/191408643-a29561da-f96a-45d7-bd1a-b30f59312ba5.png)

#### The VBA code recommended in module 2 encompasses of initializing all necessary variables for calculation, initializing array for stock tickers and making use of nested loop for calculating total volume and total returns for all tickets and all the data available.The code performed well and executed to generate a result for 2018 stock data at 0.61718 seconds.The snip of performance time is presented below:
![Module 2 performance](https://user-images.githubusercontent.com/108683284/191410027-73853948-53a3-4b81-8fb4-36f3fd58a479.png)

#### Module 2 recommended VBA Code
```
Sub AllStockAnalysis()

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
 
    '1b) Creating three output arrays.
    Dim tickerVolumes As Long
    Dim tickerStartingPrices As Double
    Dim tickerEndingPrices As Double
    
    'Get the number of rows to loop over
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
    '1a) Create a ticker Index
    For i = 0 To 11 'to LOOP over the Tickers
    
        'Setting tickerIndex to 0 before looping over the rows
        tickerIndex = tickers(i)
        tickerVolumes = 0
        
        Worksheets(yearValue).Activate
        For j = 2 To RowCount
        
                If Cells(j, 1).Value = tickerIndex Then
   
                    tickerVolumes = tickerVolumes + Cells(j, 8).Value
    
                End If
                
                
                If Cells(j, 1).Value = tickerIndex And Cells(j - 1, 1).Value <> tickerIndex Then
        
                    tickerstartingPrice = Cells(j, 6).Value
        
                End If
                
                If Cells(j, 1).Value = tickerIndex And Cells(j + 1, 1).Value <> tickerIndex Then
        
                    tickerEndingPrices = Cells(j, 6).Value
        
        
                End If
                 
        
            Next j
 Worksheets("All Stocks Analysis").Activate
 
    Cells(4 + i, 1).Value = tickerIndex 'Cells(4, 1).Value = 2018
    Cells(4 + i, 2).Value = tickerVolumes 'Cells(4, 2).Value = totalVolume
    Cells(4 + i, 3).Value = tickerEndingPrices / tickerstartingPrice - 1
    
Next i

'Formatting
    Worksheets("All Stocks Analysis").Activate
    Range("A3:C3").Font.FontStyle = "Bold"
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Range("B4:B15").NumberFormat = "#,##0"
    Range("C4:C15").NumberFormat = "0.0%"
    Columns("B").AutoFit
    Columns("C").AutoFit
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

#### 2.2 Refactored VBA code
#### The result of stock data analysis for year 2018 using refactored VBA code is presented below:
![refactored result 2018](https://user-images.githubusercontent.com/108683284/191412857-d9740203-f49f-4a61-9f6c-5ed52dca3c66.png)

#### The refactored VBA code also initializes all the necessary variables and set tickers in array. Along with tickers, the tickerVolumes, tickerStartingPrices and tickerEndingPrices are also initialized as array. Two different loops were used instead of nested loop, first to  to  set the ticker into ticker array and respective calculated  tickerVolumes,tickerStartingPrices and tickerEndingPrices were stored in array. Instead of running (11 X 3013) loops as in module 2 recommended VBA code using nested loops, refactored VBA code used first loop 11 times and second loop 3013 times for calculation. The result from refactored code presented better performance compared to code in module 2. The code execution time for refactored VBA code was 0.0859 seconds.   
![Refactored Code  performance](https://user-images.githubusercontent.com/108683284/191414337-b5226e6e-8b5c-4b59-b891-d8db3e597667.png)

#### Refactored VBA Code
```
Sub AllStocksAnalysisRefactored()
 Dim startTime As Single
    Dim endTime As Single
    
    yearValue = InputBox("What year would you like to run the analysis on?")
    
     startTime = Timer
     
    '1) Format the output sheet on All Stocks Analysis worksheet
    Worksheets("All Stocks Analysis").Activate
    
    Range("A1").Value = "All Stocks (" + yearValue + ")"
    
    '2) Create a header row
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"

    '3) Initialize array of all tickers
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
    
    Worksheets(yearValue).Activate
    
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
    Dim tickerIndex As Single
    tickerIndex = 0
    
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    
     For i = 0 To 11
    tickerVolumes(i) = 0
    
    Next i
    
    For i = 2 To RowCount
    
    tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
    
    If Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
    
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
            
            End If
            
            
             If Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
             
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
            
            tickerIndex = tickerIndex + 1
            
     End If
   Next i
   
   For i = 0 To 11
   Worksheets("All Stocks Analysis").Activate
    tickerIndex = i
    Cells(i + 4, 1).Value = tickers(tickerIndex)
    Cells(i + 4, 2).Value = tickerVolumes(tickerIndex)
    Cells(i + 4, 3).Value = tickerEndingPrices(tickerIndex) / tickerStartingPrices(tickerIndex) - 1
   Next i
   
   '9) Formatting
    Worksheets("All Stocks Analysis").Activate
    Range("A3:C3").Font.FontStyle = "Bold"
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Range("B4:B15").NumberFormat = "#,##0"
    Range("C4:C15").NumberFormat = "0.0%"
    Columns("B").AutoFit
    Columns("C").AutoFit
    

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
   MsgBox "This code ran in " & (endTime - startTime) & "seconds for the year " & (yearValue)
   
    
End Sub
```
## 3. Summary 

#### The advantages of refactoring the code is it helps reducing the complexity of the code and improve code reliability to minimize the bugs. It helps in reducing the run time of the code and make debugging simple. The major drawback of refactoring the code is sometime, refactoring can introduce a bug or errors which is difficult for the coder to identify and figure ourt the solution. 

#### In our case, the module 2 recommended VBA code, which is original code for the stock data analysis is very simple to understand because the code was prepared from scratch and all the logicals and loops used were straightforward. However, with the simple and straightforward code, the use of nested loop to calculate and return the value for total volume and stock return was slower. The refactored code is conceptualized and written in a very different way to store all calculations in an array and two seperate loops were used instead of nested loop. The use of array and seperate nested loop made the code confusing compared to original code. Despite being confusing code at a glance, the execution of the refactored code was 90% faster compared to the original code.  
