Sub StockDataModerate()

    
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
        
        
        
        
        
        
        
            
        
    Next ws
    
End Sub


