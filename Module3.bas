Attribute VB_Name = "Module3"
Sub Stock3()

'Declare and set worksheet
Dim ws As Worksheet

'Loop through all stocks for one year
For Each ws In Worksheets

'Create the column headings
ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Volume"

ws.Range("P1").Value = "Ticker"
ws.Range("Q1").Value = "Value"
ws.Range("O2").Value = "Greatest % Increase"
ws.Range("O3").Value = "Greatest % Decrease"
ws.Range("O4").Value = "Greatest Total Volume"

Dim ticker As String
ticker = " "
Dim Ticker_volume As Double
Ticker_volume = 0

Dim LastRow As Long
Dim i As Long
Dim j As Integer

'Define Lastrow of worksheet
LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

'Set new variables for prices and percent changes
Dim open_price As Double
open_price = 0
Dim close_price As Double
close_price = 0
Dim price_change As Double
price_change = 0
Dim price_change_percent As Double
price_change_percent = 0


'Do loop of current worksheet to Lastrow
For i = 2 To LastRow

'Ticker symbol output
If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
ticker = ws.Cells(i, 1).Value

'Calculate change in Price
close_price = ws.Cells(i, 6).Value
price_change_percent = close_price - open_price

'Fixing the open price equal zero problem
ElseIf open_price <> 0 Then
price_change_percent = (price_change_percent / open_price) * 100

End If

Next i

Next ws

End Sub


