Attribute VB_Name = "Module1"
Sub StockAnalysis()
    For Each ws In Worksheets
        Dim WorksheetName As String
        WorksheetName = ws.Name
        
        Dim LastRow As Long
        'Find the last row with data'
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        'MsgBox (LastRow)'
        
        Dim PercentChange As Double
        
        'Total Stock Volume - set to 0"
        Dim TotalStockVol As Long
        TotalStockVol = ws.Cells(2, 7).Value
        
        Dim TickerCount As Long
        'Set TickerCount to first row'
        TickerCount = 2
        
        Dim j As Long
        j = 2
        
        Dim GreatestVol As Double
        Dim GreatestIncrease As Double
        Dim GreatestDecrease As Double
        
        
        'Create column headers'
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        
        For i = 2 To LastRow
            'Did the ticker name change?'
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                'If yes then...'
            
                'Write Ticker in column 9'
                ws.Cells(TickerCount, 9).Value = ws.Cells(i, 1).Value
                'Calculate and write yearly change in column 10'
                ws.Cells(TickerCount, 10).Value = ws.Cells(i, 6).Value - ws.Cells(j, 3).Value
                
                'Add conditional formatting'
                If ws.Cells(TickerCount, 10).Value < 0 Then
                    'If less than 0, set colour to red'
                    ws.Cells(TickerCount, 10).Interior.ColorIndex = 3
                Else
                    'else, set colour to green'
                    ws.Cells(TickerCount, 10).Interior.ColorIndex = 4
                End If
                
                'Calculate percent change'
                If ws.Cells(j, 3).Value <> 0 Then
                    'Calculation'
                    PercentChange = ((ws.Cells(i, 6).Value - ws.Cells(j, 3).Value) / ws.Cells(j, 3).Value)
                    'Percent formating'
                    ws.Cells(TickerCount, 11).Value = Format(PercentChange, "Percent")
                Else
                    ws.Cells(TickerCount, 11).Value = Format(0, "Percent")
                End If
           
                'Calculate and record total volume in col 12'
                ws.Cells(TickerCount, 12).Value = WorksheetFunction.Sum(Range(ws.Cells(j, 7), ws.Cells(i, 7)))
                
                'Increase TickerCount by 1'
                TickerCount = TickerCount + 1
                'New start row for results of code'
                j = i + 1
            End If
            
            'Conditional formatting - percent change'
            If ws.Cells(i, 11).Value >= 0 Then
                'Colour green'
                ws.Cells(i, 11).Interior.ColorIndex = 4
            ElseIf ws.Cells(i, 11).Value < 0 Then
                'Colour - red'
                ws.Cells(i, 11).Interior.ColorIndex = 3
            End If
        Next i
        
        'Find the last row in col 9 (col 9 to col 12 stats)'
        Dim LastRowCol9 As Long
        LastRowCol9 = ws.Cells(Rows.Count, 9).End(xlUp).Row
        'MsgBox (LastRowCol9)'
        
        'Initialize summary variables to first row of data'
        GreatestVolume = ws.Cells(2, 12).Value
        GreatestIncrease = ws.Cells(2, 11).Value
        GreatestDecrease = ws.Cells(2, 11).Value
        
        'Loop through to get a summary'
        For i = 2 To LastRowCol9
            'If statement to check if value in col 12 is greater than the current greatest volume'
            If ws.Cells(i, 12).Value > GreatestVolume Then
                GreatestVolume = ws.Cells(i, 12).Value
                'update greatest volume if greater'
                ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
            Else
                GreatestVolume = GreatestVolume
            End If
            
            'If statement to check for greatest value in col 11'
            If ws.Cells(i, 11).Value > GreatestIncrease Then
                GreatestIncrease = ws.Cells(i, 11).Value
                'update greatest increase if greater'
                ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
            Else
                GreatestIncrease = GreatestIncrease
            End If
            
            'If statement to check if value is smaller in col 11'
            If ws.Cells(i, 11).Value < GreatestDecrease Then
                GreatestDecrease = ws.Cells(i, 11).Value
                'update greatest decrease if smaller'
                ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
            Else
                GreatestDecrease = GreatestDecrease
            End If
            
            'print the summary results and format'
            ws.Cells(2, 17).Value = Format(GreatestIncrease, "Percent")
            ws.Cells(3, 17).Value = Format(GreatestDecrease, "Percent")
            ws.Cells(4, 17).Value = Format(GreatestVolume, "Scientific")
        Next i
        'Adjust column width automatically'
        Worksheets(WorksheetName).Columns("A:Z").AutoFit
    Next ws
End Sub
