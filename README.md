Sub AlphabetTesting():

Dim ws As Worksheet
For Each ws In Worksheets

ws.Cells(1, 9) = "Ticker"
ws.Cells(1, 10) = "Yearly Change"
ws.Cells(1, 11) = "Percent Change"
ws.Cells(1, 12) = "Total Stock Volume"
    
    Dim summary_row As LongLong 'Integer
    Dim total_volume As LongLong 'Double
    
        
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    ticker_row = 2
    total_volume = 0
    
    For i = 2 To lastrow
        If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
            ticker = ws.Cells(i, 1).Value
            total_volume = total_volume + ws.Cells(i, 7).Value
            ws.Range("I" & ticker_row).Value = ticker
            ws.Range("L" & ticker_row).Value = total_volume
            ticker_row = ticker_row + 1
            
            total_volume = 0
            
        Else
          total_volume = total_volume + ws.Cells(i, 7).Value
        End If
    Next i
    
'Finding Yearly Change and Percent Change
        
        Dim year_open As Double
        Dim year_close As Double
        Dim change_row As LongLong
        Dim yearly_change As Double
        Dim starting_open As Double
        Dim percent_change As Double
        
        starting_open = ws.Cells(2, 3).Value
        change_row = 2
        
    For i = 2 To lastrow
        If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
            year_open = ws.Cells(i, 3).Value
            year_close = ws.Cells(i, 6).Value
            yearly_change = year_close - starting_open
            ws.Range("J" & change_row).Value = yearly_change
            
            'Finding Percent Change
            percent_change = ws.Range("J" & change_row).Value / ws.Cells(i, 3).Value
            ws.Range("K" & change_row).Value = percent_change
            ws.Range("K" & change_row).NumberFormat = "0.00%"
            change_row = change_row + 1
            
            
            'Reset starting_open variable
            starting_open = ws.Cells(i + 1, 3).Value
        End If
        
        'Add Conditional Coloring
        
        If ws.Cells(i, 10).Value > 0 Then
            ws.Cells(i, 10).Interior.ColorIndex = 4
        ElseIf ws.Cells(i, 10).Value < 0 Then
            ws.Cells(i, 10).Interior.ColorIndex = 3
        End If
    Next i
Next ws
End Sub
