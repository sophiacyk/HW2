Sub Volume()
Dim Ticker_current, Ticker_previous As String
Dim VolumeCount, TickerCount As Double
Dim ws As Worksheet
Dim LastRow As Double
Dim OpenPrice, ClosePrice, PriceChange, YearChange As Double
Dim MaxIncre, MaxDecre, MaxVolume As Double
Dim IncreTic, DecreTic, VolTic As String



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
    MaxIncre = 0
    MaxDecre = 0
    MaxVolume = 0
    

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
            
            If VolumeCount > MaxVolume Then
                MaxVolume = VolumeCount
                VolTic = Ticker_previous
            End If
     
            
          
            ClosePrice = ws.Cells(i - 1, 6)
            PriceChange = ClosePrice - OpenPrice
            
            'Output the price change
            ws.Cells(TickerCount, 13) = PriceChange
    
            
            If Not OpenPrice = 0 Then
                ws.Cells(TickerCount, 14) = PriceChange / OpenPrice
                ws.Cells(TickerCount, 14).NumberFormat = "0.00%"
                
            
                If PriceChange >= 0 Then
                    ws.Cells(TickerCount, 13).Interior.ColorIndex = 10
                    
                    
                    If ws.Cells(TickerCount, 14) > MaxIncre Then
                        MaxIncre = ws.Cells(TickerCount, 14)
                        IncreTic = ws.Cells(TickerCount, 11)
                    End If
                    
                Else
                    ws.Cells(TickerCount, 13).Interior.ColorIndex = 3
                    
                    
                    If ws.Cells(TickerCount, 14) < MaxDecre Then
                        MaxDecre = ws.Cells(TickerCount, 14)
                        DecreTic = ws.Cells(TickerCount, 11)
                    End If
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
            
            If VolumeCount > MaxVolume Then
                MaxVolume = VolumeCount
                VolTic = Ticker_current
            End If
            
            'Get the closing price and the delta
            ClosePrice = ws.Cells(i, 6)
            PriceChange = ClosePrice - OpenPrice
            ws.Cells(TickerCount, 13) = PriceChange
            
            
            If Not OpenPrice = 0 Then
            ws.Cells(TickerCount, 14) = PriceChange / OpenPrice
            ws.Cells(TickerCount, 14).NumberFormat = "0.00%"
            
                If PriceChange >= 0 Then
                    ws.Cells(TickerCount, 13).Interior.ColorIndex = 10
                    
                    If ws.Cells(TickerCount, 14) > MaxIncre Then
                        MaxIncre = ws.Cells(TickerCount, 14)
                        IncreTic = ws.Cells(TickerCount, 11)
                    End If
                Else
                    ws.Cells(TickerCount, 13).Interior.ColorIndex = 3
                    
                    
                    If ws.Cells(TickerCount, 14) < MaxDecre Then
                        MaxDecre = ws.Cells(TickerCount, 14)
                        DecreTic = ws.Cells(TickerCount, 11)
                    End If
                End If
            
            Else
            ws.Cells(TickerCount, 14) = "N/A"
          
            End If
            
            
            
        End If
   
    
    Next i
    
    ws.Range("P2") = "Greatest % Increse"
    ws.Range("P3") = "Greatest % Decrese"
    ws.Range("P4") = "Greatest Total Volume"
    ws.Range("Q1") = "Ticker"
    ws.Range("R1") = "Value"
    
    ws.Range("Q2") = IncreTic
    ws.Range("R2") = MaxIncre
    ws.Range("R2").NumberFormat = "0.00%"
    
    ws.Range("Q3") = DecreTic
    ws.Range("R3") = MaxDecre
    ws.Range("R3").NumberFormat = "0.00%"
    
    ws.Range("Q4") = VolTic
    ws.Range("R4") = MaxVolume
   
    

Next

End Sub
