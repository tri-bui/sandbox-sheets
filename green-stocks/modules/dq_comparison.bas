Attribute VB_Name = "DQComparison"
Sub DQComparison()
    ''' Analyze another stock to compare to DQ stock '''
    
    ' Display ticker list
    TickerList ' call TickerList()
    
    ' Activate "DQ Analysis" sheet
    Worksheets("DQ Analysis").Activate
    
    ' User input for 2nd ticker
    ticker = InputBox("Please enter a 2nd ticker symbol to analyze")
    
     ' Start timer
    startTime = Timer
    
    ' Create headers
    Range("A9").Value = "Ticker: " + ticker
    Cells(11, 1).Value = "Year"
    Cells(11, 2).Value = "Total Volume"
    Cells(11, 3).Value = "Starting Price"
    Cells(11, 4).Value = "Ending Price"
    Cells(11, 5).Value = "Return ($)"
    Cells(11, 6).Value = "Return (%)"
    
    ' Analyze both years
    For y = 2017 To 2018
        AnalyzeStock ticker, y ' call AnalyzeDQ(ticker, y)
    Next y
    
    ' Format sheet
    DQFormatting ' call DQFormatting()
    
    ' Timer
    endTime = Timer
    MsgBox ("The analysis took " & Round(endTime - startTime, 4) & " seconds")
    
End Sub


Sub AnalyzeStock(ticker, y)
    ''' Analyze the specified stock for the specified year '''

    ' Activate sheet to analyze
    Worksheets(CStr(y)).Activate
    
    ' Initialize variables
    Dim startPrice As Double
    Dim endPrice As Double
    volume = 0
    
    ' Rows to iterate through
    startRow = 2
    endRow = Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
    
    ' Iterate through rows
    For r = startRow To endRow
    
        ' Check if the current row is the specified stock
        If Cells(r, 1).Value = ticker Then
            
            ' Volume
            volume = volume + Cells(r, 8).Value
            
            ' Start price
            If Cells(r - 1, 1).Value <> ticker Then
                startPrice = Cells(r, 6).Value
            End If
            
            ' End price
            If Cells(r + 1, 1).Value <> ticker Then
                endPrice = Cells(r, 6).Value
                Exit For ' exit loop
            End If
            
        End If
    Next r
    
    ' Activate "DQ Analysis" sheet
    Worksheets("DQ Analysis").Activate
    
    ' Row for results
    offsetRows = 2017 - 12 ' 2017 results on row 12, 2018 results on row 13
    resultRow = y - offsetRows
    
    ' Calculate returns
    returnUsd = endPrice - startPrice
    returnPct = returnUsd / startPrice
    
    ' Add results to sheet
    Cells(resultRow, 1).Value = y
    Cells(resultRow, 2).Value = volume
    Cells(resultRow, 3).Value = startPrice
    Cells(resultRow, 4).Value = endPrice
    Cells(resultRow, 5).Value = returnUsd
    Cells(resultRow, 6).Value = returnPct
    
End Sub


Sub DQFormatting()
    '''Apply formating to "DQ Analysis" sheet '''
    
    ' Activate "DQ Analysis" sheet
    Worksheets("DQ Analysis").Activate
    
    ' Autofit columns
    Columns("A:F").AutoFit
    
    ' Title cell formatting
    With Range("A9:B9")
        .Merge ' merge cells
        .HorizontalAlignment = xlCenter ' align center
        .Font.Size = 16 ' large font
        .Font.Bold = True ' bold text
        .BorderAround ' all borders
    End With
    
    ' Header row formatting
    With Range("A11:F11")
        .Font.Bold = True ' bold text
        .Borders(xlEdgeBottom).LineStyle = xlContinuous ' bottom border
    End With
    
    ' Numeric formatting
    Range("B:B").NumberFormat = "0,000" ' total volume
    Range("C:E").NumberFormat = "$ 0.00"  ' dollar amounts
    Range("F:F").NumberFormat = "0.00 %" ' % return
    
    ' Color formatting
    For r = 12 To 13
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


Sub TickerList()
    ''' Display a list of tickers '''
    
    'Activate "DQ Analysis" sheet
    Worksheets("DQ Analysis").Activate
    
    ' Header
    With Range("K1")
        .Value = "Tickers"
        .Font.Size = 16
        .Font.Bold = True
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
    End With
    
    ' Activate "2017" sheet
    Worksheets("2017").Activate
    
    ' Initialize variables
    ticker = Cells(2, 1).Value ' current ticker
    nRows = Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row ' number of rows
    outputRow = 2 ' row to output the ticker
    
    ' Tickers
    For r = 2 To nRows
        If Cells(r + 1, 1).Value <> ticker Then
            Worksheets("DQ Analysis").Activate ' activate output sheet
            Cells(outputRow, 11).Value = ticker ' output current ticker
            outputRow = outputRow + 1 ' increment output row
            Worksheets("2017").Activate ' activate data sheet
            ticker = Cells(r + 1, 1).Value ' get next ticker
        End If
    Next r
    
    ' Border
    Worksheets("DQ Analysis").Activate ' activate output sheet
    Range("K1:K13").BorderAround
    
End Sub
