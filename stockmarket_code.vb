Sub stockmarket()

For Each ws In Worksheets

Dim i As Long
Dim ticker_counter As Double
Dim ticker As String
Dim lastrow As Long
Dim open_price As Double
Dim close_price As Double
Dim percent_change As Double
Dim total_stock As Double

ticker_counter = 1
ws.Cells(2, 9).Value = ws.Cells(2, 1).Value

open_price = ws.Cells(2, 3).Value

total_stock = 0

lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly change"
ws.Range("K1").Value = "Percent change"
ws.Range("L1").Value = "Total stock volume"

For i = 2 To lastrow
If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
ticker = ws.Cells(i + 1, 1).Value
ws.Cells(ticker_counter + 2, 9).Value = ticker
close_price = ws.Cells(i, 6).Value
ws.Cells(ticker_counter + 1, 10).Value = close_price - open_price
percent_change = ((close_price - open_price) / open_price)
ws.Cells(ticker_counter + 1, 11).Value = percent_change
ws.Cells(ticker_counter + 1, 12).Value = total_stock
open_price = ws.Cells(i + 1, 3).Value
ticker_counter = ticker_counter + 1
Else: total_stock = total_stock + ws.Cells(i, 7).Value


End If

Next i

For i = 2 To lastrow
If ws.Cells(i, 10).Value > 0 Then
    ws.Cells(i, 10).Interior.ColorIndex = 4
    ElseIf ws.Cells(i, 10).Value < 0 Then
    ws.Cells(i, 10).Interior.ColorIndex = 3
    ElseIf ws.Cells(i, 10).Value = 0 Then
    ws.Cells(i, 10).Interior.ColorIndex = 2
    
End If
    
Next i

Next ws

End Sub

'Merry Christmas, ya filthy animal! :)

