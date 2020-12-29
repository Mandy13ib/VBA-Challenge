Sub Stock_market()
For Each ws In Worksheets

Dim Ticker As String

Dim FirstPrice As Double
FirstPrice = 0

Dim LastPrice As Double
LastPrice = 0

Dim YearlyChange As Double
YearlyChange = 0

Dim Percentage As Double
Percentage = 0

Dim TotalVolume As Double
TotalVolume = 0

Dim StockVolume As Long
StockVolume = 0

'set tickerrow
Dim TickerRow As Integer


'find the value of first price and where to end
    FirstPrice = ws.Cells(2, 3).Value
   
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
     
    TickerRow = 2

'the headers for table
ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Percenatge Change"
ws.Cells(1, 12).Value = "Total Stock Volume"

For i = 2 To LastRow

'from class notes
If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

    'set value for for ticker and where to end
    Ticker = ws.Cells(i, 1).Value
    LastPrice = ws.Cells(i, 6).Value
    
'percentage
YearlyChange = LastPrice - FirstPrice
    
        If FirstPrice > 0 Then
             Percentage = YearlyChange / FirstPrice
        Else
            Percenatge = 0
        End If
   
  'color change for line j
        If (YearlyChange > 0) Then
            ws.Range("J" & TickerRow).Interior.ColorIndex = 4
        Else
            ws.Range("J" & TickerRow).Interior.ColorIndex = 3
        End If
    'find total volume
    TotalVolume = TotalVolume + ws.Cells(i, 7).Value
    
    
    ws.Range("I" & TickerRow).Value = Ticker
    ws.Range("J" & TickerRow).Value = YearlyChange
    ws.Range("K" & TickerRow).Value = Percentage
    ws.Range("L" & TickerRow).Value = TotalVolume
    ws.Range("K" & TickerRow).NumberFormat = "00.00%"
    FirstOpenPrice = ws.Cells(i + 1, 3).Value
    
    
    TickerRow = TickerRow + 1
    
    'reset
    TotalVolume = 0
    
    Else
    YearlyChange = LastPrice - FirstPrice
    TotalVolume = TotalVolume + ws.Cells(i, 7).Value
    
    End If
    
        
            
   
Next i

Next ws

End Sub