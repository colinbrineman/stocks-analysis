Sub AllStocksAnalysisRefactored()
    
    'Select output worksheet
    Worksheets("All Stocks Analysis").Activate

    'Select year to analyze
    yearValue = InputBox("What year would you like to run the analysis on?")

    'Initialize timer
    Dim startTime As Single
    Dim endTime  As Single
    startTime = Timer
    
    'Add year header to output worksheet
    Range("A1").Value = "All Stocks (" + yearValue + ")"
    
    'Add table headers to output worksheet
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"

    'Initialize ticker array
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
    
    'Select data worksheet
    Worksheets(yearValue).Activate
    
    'Get number of rows to loop over
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
    '1a) Initialize tickerIndex
    Dim tickerIndex As Integer
    tickerIndex = 0

    '1b) Initialize tickerVolumes, tickerStartingPrices, and tickerEndingPrices arrays
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    
    '2a) Create for loop to initialize tickerVolumes to zero
    For i = 0 To 11
        tickerVolumes(i) = 0
    Next i
        
    '2b) Loop over all rows in data worksheet
    Worksheets(yearValue).Activate
    For i = 2 To RowCount
    
        '3a) Increase volume for current tickerIndex
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
        '3b) Check if current row is first row with selected tickerIndex
        If Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
        End If
        
        '3c) Check if current row is last row with selected tickerIndex
        If Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value

            '3d) Increase tickerIndex
            tickerIndex = tickerIndex + 1
        End If
    
    Next i
    
    '4) Loop through arrays to output Ticker, Total Daily Volume, and Return
    For i = 0 To 11
        Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = (tickerEndingPrices(i) / tickerStartingPrices(i)) - 1
    Next i
    
    'Select output worksheet
    Worksheets("All Stocks Analysis").Activate
    
    'Format font
    Range("A3:C3").Font.FontStyle = "Bold"
    
    'Format borders
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    
    'Format numbers
    Range("B4:B15").NumberFormat = "#,##0"
    Range("C4:C15").NumberFormat = "0.0%"
    
    'Format column width
    Columns("B").AutoFit

    'Format color, conditioning on value
    dataRowStart = 4
    dataRowEnd = 15
    For i = dataRowStart To dataRowEnd
        If Cells(i, 3) > 0 Then
            Cells(i, 3).Interior.Color = vbGreen
        Else
            Cells(i, 3).Interior.Color = vbRed
        End If
    Next i
 
    'Display run time
    endTime = Timer
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)

End Sub