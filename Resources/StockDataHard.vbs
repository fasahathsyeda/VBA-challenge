Sub StockDataHard()

    
    For Each ws In Worksheets
        ws.Activate
        Dim NewTableRow, NewTableColumn As Integer
        NewTableRow = 2
        Dim ticker As String
        Dim totalStocksVolume, closePrice, openPrice As Double
        totalStocksVolume = 0
        

        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percentage Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        Cells(1, 9).ColumnWidth = 10
        Cells(1, 10).ColumnWidth = 16
        Cells(1, 11).ColumnWidth = 16
        Cells(1, 12).ColumnWidth = 16
        
        For i = 2 To LastRow
        
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
                ticker = Cells(i, 1).Value
                Cells(NewTableRow, 9).Value = ticker
                
            
                
                closePrice = Cells(i, 6).Value
                Cells(NewTableRow, 10).Value = closePrice - openPrice
               
        If Cells(NewTableRow, 10).Value > 0 Then
                    Cells(NewTableRow, 10).Interior.ColorIndex = 4
                ElseIf Cells(NewTableRow, 10).Value < 0 Then
                    Cells(NewTableRow, 10).Interior.ColorIndex = 3
                End If
                    
                
                If (openPrice <> 0) Then
                
                    Cells(NewTableRow, 11).Value = (closePrice - openPrice) / openPrice
                    'Cells(NewTableRow, 11).Style = "Percent"
                    
                    openPrice = 0
                    closePrice = 0
                Else
                    Cells(NewTableRow, 11).Value = ""
                    'Cells(NewTableRow, 11).Style = "Percent"
                End If

        
                

                'Cells(NewTableRow, 11).Style = "Percent"
                Cells(NewTableRow, 11).Value = Format(Cells(NewTableRow, 11).Value, "Percent")
                totalStocksVolume = totalStocksVolume + Cells(i, 7).Value
                Cells(NewTableRow, 12).Value = totalStocksVolume
                NewTableRow = NewTableRow + 1
                totalStocksVolume = 0

            ElseIf Cells(i - 1, 1).Value <> Cells(i, 1).Value Then
                openPrice = ws.Cells(i, 3).Value

                totalStocksVolume = totalStocksVolume + Cells(i, 7).Value

            
                
                
            Else
      
                totalStocksVolume = totalStocksVolume + Cells(i, 7).Value

            End If
                
        Next i
        
        
        Dim GreatestIncrease, GreatestDecrease, GreatestTotalVolume As Double
        Dim GreatestIncreaseTicker, GreatestDecreaseTicker, GreatestTotalVolumeTicker  As Integer
       
        Range("O1").ColumnWidth = 22
        Range("P1").Value = "Ticker"
        Range("Q1").Value = "Value"
        Range("Q1").ColumnWidth = 18
        
               
        Range("O2").Value = "Greatest % Increase"
        GreatestIncrease = WorksheetFunction.Max(ws.Range("K2:K" & (NewTableRow - 1)))
        Range("Q2").Value = Format(GreatestIncrease, "Percent")
        GreatestIncreaseTicker = WorksheetFunction.Match(Range("Q2").Value, Range("K2:K" & (NewTableRow - 1)), 0)
        Range("P2").Value = Cells(GreatestIncreaseTicker + 1, 9).Value
        
        Range("O3").Value = "Greatest % Decrease"
        GreatestDecrease = WorksheetFunction.Min(ws.Range("K2:K" & (NewTableRow - 1)))
        Range("Q3").Value = Format(GreatestDecrease, "Percent")
        GreatestDecreaseTicker = WorksheetFunction.Match(Range("Q3").Value, Range("K2:K" & (NewTableRow - 1)), 0)
        Range("P3").Value = Cells(GreatestDecreaseTicker + 1, 9).Value
        
        Range("O4").Value = "Greatest Total Volume"
        GreatestTotalVolume = WorksheetFunction.Max(ws.Range("L2:L" & (NewTableRow - 1)))
        Range("Q4").Value = GreatestTotalVolume
        GreatestTotalVolume = WorksheetFunction.Match(Range("Q4").Value, Range("L2:L" & (NewTableRow - 1)), 0)
        Range("P4").Value = Cells(GreatestTotalVolume + 1, 9).Value
        
        
        
        'Setting Ticker for Greatest % increase,Greatest % decrease and Greatest Total Volume
        'For j = 2 To NewTableRow
        
        'If Cells(j, 11).Value = Cells(2, 17) Then
         '       Cells(2, 16).Value = Cells(j, 9).Value
                
        'ElseIf Cells(j, 11).Value = Cells(3, 17).Value Then
         '       Cells(3, 16).Value = Cells(j, 9).Value
                
        'ElseIf Cells(j, 12).Value = Cells(4, 17).Value Then
         '       Cells(4, 16).Value = Cells(j, 9).Value
        'End If
        'Next j
    
    Next ws
    
End Sub



