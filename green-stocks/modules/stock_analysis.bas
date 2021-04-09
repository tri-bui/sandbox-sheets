Attribute VB_Name = "StockAnalysis"
Sub AllStockAnalysis()
    ''' Analyze all stocks for a specified year '''
    
    ' User input for year
    yr = InputBox("Please enter the year to analyze stocks for (2017 or 2018)")
    
    ' Start timer
    startTime = Timer
    
    ' Create headers in "Stock Analysis" sheet
    Worksheets("Stock Analysis").Activate ' activate sheet
    Range("A1").Value = "All Stocks (" + yr + ")"
    Cells(4, 1).Value = "Ticker"
    Cells(4, 2).Value = "Total Volume"
    Cells(4, 3).Value = "Starting Price"
    Cells(4, 4).Value = "Ending Price"
    Cells(4, 5).Value = "Return ($)"
    Cells(4, 6).Value = "Return (%)"
    
    ' Activate data sheet
    Worksheets(yr).Activate
    
    ' Initialize variables
    volume = 0
    startPrice = Cells(2, 6).Value
    Dim endPrice As Double
    Dim returnUsd As Double
    Dim returnPct As Double
    ticker = Cells(2, 1).Value ' current stock
    nRows = Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row ' number of rows in data sheet
    resultRow = 5 ' row to add results to in "Stock Analysis" sheet
    
    ' Iterate through data sheet
    For r = 2 To nRows
        volume = volume + Cells(r, 8).Value
            
        ' If current row is the last row for the current stock
        If Cells(r + 1, 1).Value <> ticker Then
        
            ' Calculate return
            endPrice = Cells(r, 6).Value
            returnUsd = endPrice - startPrice
            returnPct = returnUsd / startPrice
        
            ' Add results to "Stock Analysis" sheet
            Worksheets("Stock Analysis").Activate ' activate sheet
            Cells(resultRow, 1).Value = ticker
            Cells(resultRow, 2).Value = volume
            Cells(resultRow, 3).Value = startPrice
            Cells(resultRow, 4).Value = endPrice
            Cells(resultRow, 5).Value = returnUsd
            Cells(resultRow, 6).Value = returnPct
            
            ' Update variables
            Worksheets(yr).Activate ' activate data sheet
            ticker = Cells(r + 1, 1).Value
            startPrice = Cells(r + 1, 6).Value
            volume = 0
            resultRow = resultRow + 1
            
        End If
    Next r
    
    ' Format sheet
    StockFormatting ' call StockFormatting()
    
    ' Timer
    endTime = Timer ' end timer
    MsgBox ("The analysis took " & Round(endTime - startTime, 4) & " seconds")
    
End Sub


Sub StockFormatting()
    '''Apply formating to "Stock Analysis" sheet '''
    
    ' Activate "Stock Analysis" sheet
    Worksheets("Stock Analysis").Activate
    
    ' Resize columns
    Columns("A:F").ColumnWidth = 15
    
    ' Title cell formatting
    With Range("A1:B2")
        .Merge ' merge cells
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Size = 16
        .Font.Bold = True
        .BorderAround
    End With
    
    ' Header row formatting
    With Range("A4:F4")
        .Font.Bold = True
        .Borders(xlEdgeBottom).Weight = xlMedium
    End With
    
    ' Table borders
    With Range("A4:F16")
        .BorderAround , Weight:=xlMedium
        .Borders(xlInsideVertical).Weight = xlMedium
    End With
    
    ' Numeric formatting
    Range("B:B").NumberFormat = "0,000" ' total volume
    Range("C:E").NumberFormat = "$ 0.00"  ' dollar amounts
    Range("F:F").NumberFormat = "0.00 %" ' % return
    
    ' Color formatting
    nStocks = 12
    For r = 5 To 4 + nStocks
        If Cells(r, 6).Value > 0 Then ' positive return
            For c = 1 To 6
                Cells(r, c).Interior.color = vbGreen
            Next c
        ElseIf Cells(r, 6).Value < 0 Then ' negative return
            For c = 1 To 6
                Cells(r, c).Interior.color = vbRed
            Next c
        Else ' no return
            For c = 1 To 6
                Cells(r, c).Interior.color = xlNone
            Next c
        End If
    Next r
End Sub
