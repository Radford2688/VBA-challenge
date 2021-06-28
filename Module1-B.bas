Attribute VB_Name = "Module1"

Sub StockData()

' Definitions
Dim i As Long
Dim counter As Long
Dim ticker As String
Dim vol As Double
Dim open_price As Double
Dim close_price As Double
Dim yearly_change As Double
Dim percent_change As Double
Dim total_volume As Double
Dim greatest_percent_increase As String
Dim greatest_percent_decrease As String

' Name Columns
    Cells(1, 9).Value = "ticker"
    Cells(1, 10).Value = "yearly change"
    Cells(1, 11).Value = "percent change"
    Cells(1, 12).Value = "total stock volume"

' Extract Ticker Symbols

    ticker = Cells(2, 1).Value
    Cells(2, 9).Value = ticker
  
    
    counter = 2
    For i = 2 To ws.
        If (Cells(i, 1).Value <> ticker) Then
            counter = counter + 1
            ticker = Cells(i, 1).Value
            Cells(counter, 9).Value = ticker
            
            
 '  Yearly Change is Close_Price minus Open_Price
 '  Open Price = i
 '  Close Price = i - 1
 
 ' loop over each worksheet in the workbook
For Each ws In Worksheets
ws.Activate


' Initialize variables for each worksheet.
ticker = ""
yearly_change = 0
opening_price = 0
percent_change = 0
total_stock_volume = 0
    

 'Calculate change in Price
If opening_price = 0 Then
    opening_price = Cells(i, 3).Value
End If

' Get the end of year closing price for ticker
closing_price = Cells(i, 6)

' Add Yearly Change to Appropriate Cell

Cells(number_tickers + 1, 10).Value = yearly_change

' If change value is greater than 0, shade cell green.
If yearly_change > 0 Then
Cells(number_tickers + 1, 10).Interior.ColorIndex = 4

' If change value is less than 0, shade cell red.
ElseIf yearly_change < 0 Then
Cells(number_tickers + 1, 10).Interior.ColorIndex = 3

' If  value is 0, shade cell yellow.
Else
Cells(number_tickers + 1, 10).Interior.ColorIndex = 6
End If

' Calculate percent change value for ticker.
If opening_price = 0 Then
percent_change = 0
Else
percent_change = (yearly_change / opening_price)
End If

' Set opening price back to 0 when we get to a different ticker in the list.
opening_price = 0
             
 ' Add total stock volume value to the appropriate cell in each worksheet.
 Cells(number_tickers + 1, 12).Value = total_stock_volume
             
 ' Set total stock volume back to 0 when we get to a different ticker in the list.
 totalk_volume = 0
 
 End If
 
 Next i
  
  ' Display Greatest Increase, Decrease, and
  Range("O2").Value = "Greatest % Increase"
  Range("O3").Value = "Greatest % Decrease"
  Range("O4").Value = "Greatest Total Volume"
  Range("P1").Value = "Ticker"
  Range("Q1").Value = "Value"
    
    ' Get the last row
    lastRowState = ws.Cells(Rows.Count, "I").End(xlUp).Row
    
    ' Initialize variables and set values of variables initially to the first row in the list.
    greatest_percent_increase = Cells(2, 11).Value
    greatest_percent_increase_ticker = Cells(2, 9).Value
    greatest_percent_decrease = Cells(2, 11).Value
    greatest_percent_decrease_ticker = Cells(2, 9).Value
    greatest_stock_volume = Cells(2, 12).Value
    greatest_stock_volume_ticker = Cells(2, 9).Value
           
            
            Debug.Print (i)
            
        End If
Next i

    
End Sub
