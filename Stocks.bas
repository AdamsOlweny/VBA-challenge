Sub Stocks():

    For Each ws In Worksheets
    
        Dim TabName As String
        'Last row column A
        Dim LastCellColumnA As Long
        'last row column I
        Dim LastRowI As Long
        'Index counter for Ticker
        Dim TickCount As Long
        'Percentage change
        Dim PercentageChange As Double
        'Variable for greatest increase
        Dim GreatestPercentageIncrease As Double
        'Greatest percentage decrease
        Dim GreatestPercentageDecrease As Double
        'Greatest total volume
        Dim GreatestTotalVolume As Double
        'Position on row
        Dim i As Long
        'Ticker block
        Dim j As Long
        
        'Get TabName
        TabName = ws.Name
        
        'Assign Headers
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        
        'Point Ticker to first row
        TickCount = 2
        
        'Start at second row
        j = 2
        
        'Last  cell in column A
        LastCellColumnA = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
            'Loop through all rows
            For i = 2 To LastCellColumnA
            
                'Check if ticker name changed
                If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                
                'Write ticker in column I - 9
                ws.Cells(TickCount, 9).Value = ws.Cells(i, 1).Value
                
                'Calculate Yearly Change in column J - 10
                ws.Cells(TickCount, 10).Value = ws.Cells(i, 6).Value - ws.Cells(j, 3).Value
                
                    'Color Coded Conditional Formating
                    If ws.Cells(TickCount, 10).Value < 0 Then
                
                    'Set cell background to red (index 3)
                    ws.Cells(TickCount, 10).Interior.ColorIndex = 3
                
                    Else
                
                    'Set cell background color to green (index 4)
                    ws.Cells(TickCount, 10).Interior.ColorIndex = 4
                
                    End If
                    
                    'Calculate percentage change in column K - 11
                    If ws.Cells(j, 3).Value <> 0 Then
                    PercentageChange = ((ws.Cells(i, 6).Value - ws.Cells(j, 3).Value) / ws.Cells(j, 3).Value)
                    
                    'Percent formating
                    ws.Cells(TickCount, 11).Value = Format(PercentageChange, "Percent")
                    
                    Else
                    
                    ws.Cells(TickCount, 11).Value = Format(0, "Percent")
                    
                    End If
                    
                'Calculate Total Volume in column L - 12
                ws.Cells(TickCount, 12).Value = WorksheetFunction.Sum(Range(ws.Cells(j, 7), ws.Cells(i, 7)))
                
                'Increase TickCount by 1
                TickCount = TickCount + 1
                
                'Set new start row of the ticker block
                j = i + 1
                
                End If
            
            Next i
            
        'Find last non-blank cell in column I
        LastRowI = ws.Cells(Rows.Count, 9).End(xlUp).Row
        'MsgBox ("Last row in column I is " & LastRowI)
        
        'Prepare for summary
        GreatestTotalVolume = ws.Cells(2, 12).Value
        GreatestPercentageIncrease = ws.Cells(2, 11).Value
        GreatestPercentageDecrease = ws.Cells(2, 11).Value
        
            'Summary loop
            For i = 2 To LastRowI
            
                'Greatest Total Volume
                'Populate cells with new value if next value is larger
                If ws.Cells(i, 12).Value > GreatestTotalVolume Then
                GreatestTotalVolume = ws.Cells(i, 12).Value
                ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
                
                Else
                
                GreatestTotalVolume = GreatestTotalVolume
                
                End If
                
                'Greatest Increase
                'Populate cells if next value is larger
                If ws.Cells(i, 11).Value > GreatestPercentageIncrease Then
                GreatestPercentageIncrease = ws.Cells(i, 11).Value
                ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
                
                Else
                
                GreatestPercentageIncrease = GreatestPercentageIncrease
                
                End If
                
                'Greatest Decrease
                'Populate cells if next value is smaller
                If ws.Cells(i, 11).Value < GreatestPercentageDecrease Then
                GreatestPercentageDecrease = ws.Cells(i, 11).Value
                ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
                
                Else
                
                GreatestPercentageDecrease = GreatestPercentageDecrease
                
                End If
                  
            'Summarize results in cells
            ws.Cells(2, 17).Value = Format(GreatestPercentageIncrease, "Percent")
            ws.Cells(3, 17).Value = Format(GreatestPercentageDecrease, "Percent")
            ws.Cells(4, 17).Value = Format(GreatestTotalVolume, "Scientific")
            
            Next i
            
        'Flex column
        Worksheets(TabName).Columns("A:Z").AutoFit
            
    Next ws
        
End Sub