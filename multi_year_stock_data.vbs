Sub StockCount()
        
For Each ws In Worksheets

  'Define all variables
    Dim tickerName As String
    Dim prevTicker As String
    Dim tickerTotal As Double
    tickerTotal = 0
   
   'Establish rows
        Dim Sumrow As Double
        Sumrow = 2
        Dim LastRow As Double

  
    ws.Range("J1").Value = "Ticker"
    ws.Range("K1").Value = "Total"

       
    
            'Select last row
            LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
            'Debug.Print (LastRow)
    
            prevTicker = ws.Cells(2, 1).Value
    
            'For loop
        For i = 2 To LastRow + 1
            tickerName = ws.Cells(i, 1).Value
            If tickerName = prevTicker Then
    
                'Add total ticker
                tickerTotal = tickerTotal + ws.Cells(i, 7).Value
            Else
                'Print the ticker name
                ws.Range("J1").Cells(Sumrow, 1).Value = prevTicker
                'Print the ticker total
                ws.Range("J1").Cells(Sumrow, 2).Value = tickerTotal
                'Add one to the summary table row
                Sumrow = Sumrow + 1
                'Start new ticker total
                tickerTotal = ws.Cells(i, 7).Value
                prevTicker = tickerName
        End If
    Next i
Next ws

End Sub
