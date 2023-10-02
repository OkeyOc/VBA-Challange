Attribute VB_Name = "Module1"
Sub Alphabetical_Test()

'declare variables
Dim Ticker As String
Dim total_stock_volume As Double
Dim percent_change As Double
Dim Ticker_Row As Integer
Dim ws As Worksheet
Dim yearly_change As Double
Dim Greatest_Increase As Double
Dim Greatest_decrease As Double
Dim Greatest_Stock_total_volume As Double
Dim open_price As Double
Dim lastrowA As Long


'loop through all Worksheets
'loop through column A
On Error Resume Next
For Each Cell In ws.Range("A2:A" & lastrowA)

For Each ws In Worksheets

j = 2
total_stock_volume = 0
Ticker_Row = 2
open_price = ws.Cells(2, 3).Value


Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly_Change"
Cells(1, 11).Value = "Pecent_change"
Cells(1, 12).Value = "Total_stock_volume"

lastrowA = ws.Cells(Rows.Count, "A").End(xlUp).Row

'loop to search through symbols

For i = 2 To lastrowA

'If Ticker changes then get Ticket value
If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then

'print Values in column i
Ticker = ws.Cells(i, 1).Value
total_stock_volume = totalstock + ws.Cells(i, 7).Value

ws.Range("i" & Ticker_Row).Value = Ticker
ws.Range("L" & Ticker_Row).Value = total_stock_volume

'calculate yearly_change and Percent_change, as 0 value

yearly_change = ws.Cells(i, 6) - open_price
ws.Range("j" & Ticker_Row).Value = yearly_change


If open_price = 0 Then
 percent_change = 0

Else

percent_change = yearly_change / open_price

End If

ws.Range("K" & Ticker_Row).Value = percent_change

'reset variable to chane row,total volume for ticker and change the new opn_Price for ticker
Ticker_Row = Tiicker_row + 1
total_stock_volume = 0
open_price = ws.Cells(1 + i, 3)

Else

total_stock_volume = total_stock_volume + ws.Cells(i, 7).Value

End If

Next i

'conditional formatting for Positive and negative values.

lr_yearly_change = ws.Cells(Rows.Count, 10).End(x1up).Row

For r = 2 To lr_yearly_change

If ws.Cells(r, 10).Value < 0 Then
ws.Cells(r, 10).Interior.ColorIndex = 4

ElseIf ws.Cells(r, 10).Value > 0 Then
ws.Cells(r, 10).Interior.ColorIndex = 3

End If

Next r
'correct the percent_change to to %

For k = 2 To lr_yearly - Change
ws.Range("K2:K" & lr_yearly_change).NumberFormat = "0.00%"

Next k


'name the greatest increase and decrease value and ticker

ws.Cells(2, 15).Value = "GreatestIncrease%"
ws.Cells(3, 15).Value = "GreatestDecrease%"
ws.Cells(4, 15).Value = "GreatestStockVolume"
ws.Cells(1, 16).Value = "ticker"
ws.Cells(1, 17).Value = "value"


'loop thorugh percen_change and total_stock_vloume

lr_percent_change = ws.Cells(Rows.Count, 11).End(x1up).Row

For s = 2 To lr_perrcent_change


'determine Max and Min percent change

If ws.Cells(s, 11).Value = Application.worksheetfuntion.Min(ws.Range("k2:K" & lr_percent_change)) Then


ws.Cells(5, 17).Value = ws.Cells(s, 11).Value
ws.Cells(5, 16).Value = ws.Cells(s, 9).Value
ws.Range("q2").numbformat = "0.00%"

'determin max total stock volume

ElseIf ws.Cells(s, 12).Value = Application.WorksheetFunction.Max(ws.Range("l2:L" & lr_percent_change)) Then

ws.Cells(4, 17).Value = ws.Cells(s, 12).Value
ws.Cells(4, 16).Value = ws.Cells(s, 9).Value

End If

Next s

End Sub
