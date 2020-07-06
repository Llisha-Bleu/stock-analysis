Sub AllStocksAnalysisRefactored()

   'Add the `startTime` and `endTime` variables as Single data types
    Dim startTime As Single
    Dim endTime As Single

    yearValue = InputBox("What year would you like to run the analysis on?")
    
        'Set the startTime variable to the `Timer` function.
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
    
    '4a) Activate data worksheet
    Worksheets(yearValue).Activate
    
    '4b) Get the number of rows to loop over
    'tells spreadsheet to run the loop through the end of the full data
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
    '5a) Create a ticker Index and set it to zero before iterating all over the rows
    Dim tickerIndex As Integer
    tickerIndex = 0

    '5b) Create three output arrays
    'Created Dim variables as we are evaluating 12 tickers
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    
    '6a) A `For` was created to initialize ticker volumes to zero
    '1st loop ID variables for specific tickers listed in the array above
    
     For j = 0 To 11
        
            tickerVolumes(0) = 0
        
        Next j
        
    '6b) Create a `For` loop iterator "i" that we will use to loop over all the rows one by one. Iterator is started at 2, to skip the header
    '2nd loop for nested loop using "i" as variable
        
        For i = 2 To RowCount
    
            '7a) Increases the current `tickerVolumes` variable and adds the ticker volume for current stock ticker. Here the `tickerIndex variable is used as the Index
            If Cells(i, 1).Value = tickers(tickerIndex) Then
            
                tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
            
            End If
        
            '7b) Check if the current row is the first row with the selected tickerIndex.
            If Cells(i - 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
                
                tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
            
            End If
        
            '7c) Check if the current row is the last row with the selected ticker
            If Cells(i + 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
                
                tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
             
             '7d Increase the tickerIndex if the next row's ticker doesnt match the previous row's ticker
                tickerIndex = tickerIndex + 1
                
            End If
        
    Next i
    
    '8) Loop through your arrays `(tickers, tickerVolumes, tickerStartingPrices, and tickerEndingPrices)` to output the Ticker, Total Daily Volume, and Return.
    'Here we want to ouput what we learned from our analysis so we switch back to our output worksheet. We assigned the values in our data table using the cells method
    For i = 0 To 11
    
    Worksheets("All Stocks Analysis").Activate
    
        'Where to list output
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
    
    'closing i second as it was the initial loop opened.
    Next i
    
    '9) Formatting
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
    
        'Set the endTime variable to the `Timer` function. A message box statement was created to tell how long the code took to run by subtracting the `startTime` from the `endTime` for the analysis
        endTime = Timer
        MsgBox "This code ran in " & (endTime - startTime) & " second for the year " & (yearValue)

End Sub