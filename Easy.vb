Sub Volume()
Dim Ticker_current, Ticker_previous As String
Dim VolumeCount, TickerCount As Long
Dim ws As Worksheet
Dim LastRow As Long


Set ws = ActiveSheet

LastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
MsgBox (LastRow)

Ticker_previous = Cells(2, 1)
VolumeCount = Cells(2, 7)
TickerCount = 2  ' for output the ticker name

For i = 3 To LastRow
Ticker_current = Cells(i, 1)
    If Ticker_current = Ticker_previous Then
    VolumeCount = VolumeCount + Cells(i, 7)
    Else
    Cells(TickerCount, 11) = Ticker_previous
    Cells(TickerCount, 12) = VolumeCount
    Ticker_previous = Cells(i, 1)
    VolumeCount = Cells(i, 7)
    TickerCount = TickerCount + 1
    End If

    If i = LastRow Then
    Cells(TickerCount, 11) = Ticker_current
    Cells(TickerCount, 12) = VolumeCount
    End If
    
Next i

End Sub
