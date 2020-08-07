Sub AllStocksAnalysisRefactored()

    'Variables to store start and end time of execution
    Dim startTime As Single
    Dim endTime  As Single

    'Store the year value entered for analysis
    yearValue = InputBox("What year would you like to run the analysis on?")

    'Store the timer value at the beginning of execution
    startTime = Timer
    
    'Format the output sheet on All Stocks Analysis worksheet
    Worksheets("All Stocks Analysis").Activate
    
    'Write the heading
    Range("A1").Value = "All Stocks (" + yearValue + ")"
    
    'Create a header row
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"

    'Initialize array of all tickers
    Dim tickers(12) As String
    
    'Assign the tickers
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
    Dim tickerIndex As Integer

    'Initialize the tickerIndex to zero
    tickerIndex = 0

    '1b) Create three output arrays
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    
    '2a) Initialize ticker volumes to zero
    For i = 0 To 11
        tickerVolumes(i) = 0
    Next i

    '2b) loop over all the rows
    For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        If (Cells(i - 1, "A").Value <> tickers(tickerIndex)) Then
            tickerStartingPrices(tickerIndex) = Cells(i, "F")
        End If
        
        '3c) check if the current row is the last row with the selected ticker
        If (Cells(i + 1, "A").Value <> tickers(tickerIndex)) Then
            tickerEndingPrices(tickerIndex) = Cells(i, "F")
            
            '3d Increase the tickerIndex.
            tickerIndex = tickerIndex + 1
        End If
    
    Next i
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For tickerIndex = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        Cells(tickerIndex + 4, "A").Value = tickers(tickerIndex)
        Cells(tickerIndex + 4, "B").Value = tickerVolumes(tickerIndex)
        Cells(tickerIndex + 4, "C").Value = tickerEndingPrices(tickerIndex) / tickerStartingPrices(tickerIndex) - 1
        
    Next tickerIndex
    
    'Formatting
    Worksheets("All Stocks Analysis").Activate
    Range("A3:C3").Font.FontStyle = "Bold"
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Range("B4:B15").NumberFormat = "#,##0"
    Range("C4:C15").NumberFormat = "0.0%"
    Columns("B").AutoFit

    'Initialize the start and end row to color code the 'Return' column
    dataRowStart = 4
    dataRowEnd = 15

    For i = dataRowStart To dataRowEnd
        
        If Cells(i, 3) > 0 Then
            
            'Format the postive cell with positive value to Green
            Cells(i, 3).Interior.Color = vbGreen
            
        Else
        
            'Format the postive cell with negative or zero value to Red
            Cells(i, 3).Interior.Color = vbRed
            
        End If
        
    Next i
 
    'Store the end time of execution
    endTime = Timer
    
    'Display the time taken to run the analysis
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)

End Sub
