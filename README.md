# VBA-challenge

Sub MultiYrStock()

'Declare variables
Dim ws As Worksheet
Dim ticker_symbol As String
Dim yearly_price_change As Double
Dim yearly_percent_change As Double
Dim total_stock_volume As LongLong
Dim open_price As Double
Dim close_price As Double
Dim Summary_Table_Row As Integer

For Each ws In Worksheets

'Set counters
Summary_Table_Row = 2
total_stock_volume = 0
yearly_price_change = 0
yearly_percent_change = 0
open_price = 0
close_price = 0

ws.Range(“J1”).Value = ticker_symbol
ws.Range(“K1).Value = yearly_price_change
ws.Range(“L1”).Value = yearly_percent_change
ws.Range(“M1”).Value = total_stock_volume

'Count to the last row used in the first column; ticker symbol
lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

'Scan until the ticker symbol changes; while scanning sum the total stock volume, calculate the yearly price change and the yearly percentage change and move totals to the summary table

For i = 2 To lastrow
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    total_stock_volume = total_stock_volume + ws.Cells(i, 7).Value
    ticker_symbol = ws.Cells(i, 1).Value
    yearly_price_change = ws.Cells(i, 6).Value - open_price
    yearly_percent_change = (yearly_price_change / open_price) * 100

    ws.Cells(Summary_Table_Row, 10).Value = ticker_symbol
    ws.Cells(Summary_Table_Row, 13).Value = total_stock_volume
    ws.Cells(Summary_Table_Row, 11).Value = yearly_price_change
    ws.Cells(Summary_Table_Row, 12).Value = yearly_percent_change

'If percent change is positive then green if negative then red
    If ws.Cells(Summary_Table_Row, 12).Value >= 0 Then
    ws.Cells(Summary_Table_Row, 12).Interior.ColorIndex = 4
    ElseIf ws.Cells(Summary_Table_Row, 12).Value < 0 Then
    ws.Cells(Summary_Table_Row, 12).Interior.ColorIndex = 3
    End If

'Reset the counters so once the ticker symbols aren’t equal and it does the above calculations it starts over at 0 and moves to the next row of the summary table
    Summary_Table_Row = Summary_Table_Row + 1
    total_stock_volume = 0
    yearly_price_change = 0
    yearly_percent_change = 0

'If the ticker symbols are equal sum the total stock volume
    
Else
    total_stock_volume = total_stock_volume + ws.Cells(i, 7).Value
    If ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then
    open_price = ws.Cells(i, 3).Value

Else
    End If

End If

Next i
Next ws

End Sub
