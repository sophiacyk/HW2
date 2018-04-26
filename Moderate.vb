Sub Volume()
Dim Ticker_current, Ticker_previous As String
Dim VolumeCount, TickerCount As Double
Dim ws As Worksheet
Dim LastRow As Double
Dim OpenPrice, ClosePrice, PriceChange, YearChange As Double




For Each ws In Worksheets


    LastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    Ticker_previous = ws.Cells(2, 1)
    VolumeCount = ws.Cells(2, 7)
    TickerCount = 2  ' for output the ticker name

    ws.Cells(1, 11) = "Ticker"
    ws.Cells(1, 12) = "Total volume"
    ws.Cells(1, 13) = "Yearly price change"
    ws.Cells(1, 14) = "Yearly change rate"
    
    
    OpenPrice = ws.Cells(2, 3)

    For i = 3 To LastRow
        Ticker_current = ws.Cells(i, 1)
        
        
        
        If Ticker_current = Ticker_previous Then
        VolumeCount = VolumeCount + ws.Cells(i, 7)
        
            If OpenPrice = 0 Then
                OpenPrice = ws.Cells(i, 3)
            End If
        
        Else
            ws.Cells(TickerCount, 11) = Ticker_previous
            ws.Cells(TickerCount, 12) = VolumeCount
     
            
          
            ClosePrice = ws.Cells(i - 1, 6)
            PriceChange = ClosePrice - OpenPrice
            'Output the price change
            ws.Cells(TickerCount, 13) = PriceChange
    
            
            If Not OpenPrice = 0 Then
                ws.Cells(TickerCount, 14) = PriceChange / OpenPric
                ws.Cells(TickerCount, 14).NumberFormat = "0.00%"
            
                If PriceChange >= 0 Then
                    ws.Cells(TickerCount, 13).Interior.ColorIndex = 10
                    ws.Cells(TickerCount, 14).Interior.ColorIndex = 10
                Else
                    ws.Cells(TickerCount, 13).Interior.ColorIndex = 3
                    ws.Cells(TickerCount, 14).Interior.ColorIndex = 3
                End If
            
            Else
                ws.Cells(TickerCount, 14) = "N/A"
            
            End If
            
            Ticker_previous = Ticker_current
            VolumeCount = ws.Cells(i, 7)
            TickerCount = TickerCount + 1
            
            'Reset open price
            OpenPrice = ws.Cells(i, 3)
            
        End If
    
        If i = LastRow Then
            ws.Cells(TickerCount, 11) = Ticker_current
            ws.Cells(TickerCount, 12) = VolumeCount
            
            'Get the closing price and the delta
            ClosePrice = ws.Cells(i, 6)
            PriceChange = ClosePrice - OpenPrice
            ws.Cells(TickerCount, 13) = PriceChange
            
            
            If Not OpenPrice = 0 Then
            ws.Cells(TickerCount, 14) = PriceChange / OpenPrice
            ws.Cells(TickerCount, 14).NumberFormat = "0.00%"
            
                If PriceChange >= 0 Then
                    ws.Cells(TickerCount, 13).Interior.ColorIndex = 10
                    ws.Cells(TickerCount, 14).Interior.ColorIndex = 10
                Else
                    ws.Cells(TickerCount, 13).Interior.ColorIndex = 3
                    ws.Cells(TickerCount, 14).Interior.ColorIndex = 3
                End If
            
            Else
            ws.Cells(TickerCount, 14) = "N/A"
          
            End If
            
            
            
        End If
   
    
    Next i

Next

End Sub
