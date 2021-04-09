Attribute VB_Name = "DQAnalysis"
Sub DQAnalysis()
    ''' Analyze DQ stock for both years '''
    
    ' Start timer
    startTime = Timer
    
    ' Activate "DQ Analysis" sheet
    Worksheets("DQ Analysis").Activate
    
    ' Create headers
    Range("A1").Value = "Ticker: DQ"
    Cells(3, 1).Value = "Year"
    Cells(3, 2).Value = "Total Volume"
    Cells(3, 3).Value = "Starting Price"
    Cells(3, 4).Value = "Ending Price"
    Cells(3, 5).Value = "Return ($)"
    Cells(3, 6).Value = "Return (%)"
    
    ' Analyze both years
    For y = 2017 To 2018
        AnalyzeDQ y ' call AnalyzeDQ(y)
    Next y
    
    ' Format sheet
    DQFormatting ' call DQFormatting()
    
    ' Timer
    endTime = Timer
    MsgBox ("The analysis took " & Round(endTime - startTime, 4) & " seconds")
    
End Sub


Sub AnalyzeDQ(y)
    ''' Analyze DQ for the specified year '''

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
    
        ' Check if the current row is DQ stock
        If Cells(r, 1).Value = "DQ" Then
            
            ' Volume
            volume = volume + Cells(r, 8).Value
            
            ' Start price
            If Cells(r - 1, 1).Value <> "DQ" Then
                startPrice = Cells(r, 6).Value
            End If
            
            ' End price
            If Cells(r + 1, 1).Value <> "DQ" Then
                endPrice = Cells(r, 6).Value
                Exit For ' exit loop
            End If
            
        End If
    Next r
    
    ' Activate "DQ Analysis" sheet
    Worksheets("DQ Analysis").Activate
    
    ' Row for results
    offsetRows = 2017 - 4 ' 2017 results on row 4, 2018 results on row 5
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
    With Range("A1:B1")
        .Merge ' merge cells
        .HorizontalAlignment = xlCenter ' align center
        .Font.Size = 16 ' large font
        .Font.Bold = True ' bold text
        .BorderAround ' all borders
    End With
    
    ' Header row formatting
    With Range("A3:F3")
        .Font.Bold = True ' bold text
        .Borders(xlEdgeBottom).LineStyle = xlContinuous ' bottom border
    End With
    
    ' Numeric formatting
    Range("B:B").NumberFormat = "0,000" ' total volume
    Range("C:E").NumberFormat = "$ 0.00"  ' dollar amounts
    Range("F:F").NumberFormat = "0.00 %" ' % return
    
    ' Color formatting
    For r = 4 To 5
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
