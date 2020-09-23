Sub StockData()

Dim ws As Worksheet

'this will add headers

For Each ws In Worksheets

    ws.Range("I1").Value = "Ticker Symbol"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"

Dim FirstRow As Integer
Dim TStockVol As Double

    FirstRow = 2
    previous_i = 1
    TStockVol = 0

    EndRow = ws.Cells(Rows.Count, "A").End(xlUp).Row

Dim TickerSym As String
Dim YrOpen As Double
Dim YrClose As Double
    
        For i = 2 To EndRow

            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

            TickerSym = ws.Cells(i, 1).Value
            previous_i = previous_i + 1
            YrOpen = ws.Cells(previous_i, 3).Value
            YrClose = ws.Cells(i, 6).Value

Dim YrChange As Double
Dim PerChange As Double

            For j = previous_i To i
                TStockVol = TStockVol + ws.Cells(j, 7).Value
            Next j

            If YrOpen = 0 Then
                PerChange = YrClose

            Else
                YrChange = YrClose - YrOpen
                PerChange = YrChange / YrOpen

            End If
    
    'move data to new columns

            ws.Cells(FirstRow, 9).Value = TickerSym
            ws.Cells(FirstRow, 10).Value = YrChange
            ws.Cells(FirstRow, 11).Value = PerChange
            ws.Cells(FirstRow, 11).NumberFormat = "0.00%"  'change to percentage
            ws.Cells(FirstRow, 12).Value = TStockVol

            FirstRow = FirstRow + 1

            TStockVol = 0
            YrChange = 0
            PerChange = 0

            previous_i = i

        End If

    Next i
    
'highlight positivechange in green and negative change in red

    jEndRow = ws.Cells(Rows.Count, "J").End(xlUp).Row

        For j = 2 To jEndRow

            If ws.Cells(j, 10) > 0 Then
                ws.Cells(j, 10).Interior.ColorIndex = 4

            Else
                ws.Cells(j, 10).Interior.ColorIndex = 3
            
            End If

        Next j


Next ws

End Sub