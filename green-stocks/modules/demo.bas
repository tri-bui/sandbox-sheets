Attribute VB_Name = "Demo"
Sub MacroCheck()
    ''' Check that macros are working '''
    
    ' Store message in variable
    Dim msg As String
    msg = "Hello World!"
    MsgBox (msg)
End Sub


Sub Checkerboard()
    ''' Create a checkerboard in the "Demo" sheet
    
    ' Activate "Demo" sheet
    Worksheets("Demo").Activate
    
    ' Board
    For r = 1 To 8 ' rows
        For c = 1 To 8 ' cols
            If r Mod 2 = c Mod 2 Then
                Cells(r, c).Interior.color = vbBlack
            Else
                Cells(r, c).Interior.color = vbRed
            End If
        Next c
    Next r
    
    ' Row and col sizing
    Rows("1:8").RowHeight = 50
    Columns("A:H").ColumnWidth = 10
    
End Sub
