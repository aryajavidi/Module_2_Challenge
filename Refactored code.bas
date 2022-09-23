Attribute VB_Name = "Module1"
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
    
    '1a) Create a ticker index
    Dim tickerIndex As Integer
    
    tickerIndex = 0
    
    '1b) Create three output arrays
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    
   ''2a) Create a for loop to initialize the tickerVolumes to zero.
    For tickerIndex = 0 To 11
    
    'Initiate tickers volume
    tickerVolumes(tickerIndex) = 0
    
    'Activate worksheet
    Worksheets(yearValue).Activate
    
        '2b) Loop over all the rows in the spreadsheet.
        For i = 2 To RowCount
        
            '3a) Increase volume for current ticker.
            tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
            
               '3b) Check if the current row is the first row with the selected tickerIndex.
            If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
            
                'if it is the first row for current ticker, then set starting price.
                tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
                
            'End If
            End If
            
        '3c) Check if the current row is the last row with the current ticker.
            If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
           
                tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
                
            'End if
            End If
            
        '3d) Increase the tickerIndex.
            If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
               
                tickerIndex = tickerIndex + 1
            'End If
            End If
        Next i
    Next tickerIndex
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
    
        'Activate Output Worksheet
        Worksheets("All Stocks Analysis").Activate
        'Ticker Row Label
        Cells(4 + i, 1).Value = tickers(i)
        'Sum of Volume
        Cells(4 + i, 2).Value = tickerVolumes(i)
        'ReturnValue
        Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
        
        'Formatting
    Worksheets("All Stocks Analysis").Activate
    Range("A3:C3").Font.FontStyle = "Bold"
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Range("B4:B15").NumberFormat = "#,##0"
    Range("C4:C15").NumberFormat = "0.0%"
    Columns("B").AutoFit
        
    dataRowStart = 4
    dataRowEnd = 15
    
    Next i
    'timer
     endTime = Timer
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)

    
    End Sub
    
    
    

Sub Clearcells()

    'clear the cells on the sheet
    
    Cells.Clear

End Sub


