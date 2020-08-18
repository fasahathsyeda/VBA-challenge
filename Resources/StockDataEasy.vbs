Sub StockDataEasy()

    
    For Each ws In Worksheets
        ws.Activate
        Dim NewTableRow, NewTableColumn As Integer
        NewTableRow = 2
        Dim ticker As String
        Dim totalStocksVolume As Double
        totalStocksVolume = 0

        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Total Stock Volume"
        
        For i = 2 To LastRow
        
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
                ticker = Cells(i, 1).Value
                totalStocksVolume = totalStocksVolume + Cells(i, 3).Value
                Range("I" & NewTableRow).Value = ticker
                Range("J" & NewTableRow).Value = totalStocksVolume
                NewTableRow = NewTableRow + 1
                totalStocksVolume = 0

            Else
      
                totalStocksVolume = totalStocksVolume + Cells(i, 7).Value

            End If
                
        Next i
        
        
        
            
        
    Next ws
    
End Sub
