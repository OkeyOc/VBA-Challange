Attribute VB_Name = "Module1"
Sub Alphabet_Test()

'declare variables
Dim Ticker As String
Dim total_stock_volume As Double
Dim percent_change As Double
Dim Ticker_Row As Integer
Dim ws As Worksheet
Dim yearly_change As Double
Dim Greatest_Increase As Double
Dim Greatest_decrease As Double
Dim Greatest_Stock_volume As Double
Dim J As Double
Dim lastrowA As Long
Dim open_Price As Double
Dim Close_price As Double
Dim cells As Range
Dim s As Double


'loop through all Worksheets
'loop through column A
On Error Resume Next
For Each Cell In ws.Range("A2:A" & lastrowA)

For Each ws In Worksheets

J = 2
total_stock_volume = 0
Ticker_Row = 2
open_Price = ws.cells(2, 3).Value


ws.cells(1, 9).Value = "Ticker"
ws.cells(1, 10).Value = "Yearly_Change"
ws.cells(1, 11).Value = "Pecent_change"
ws.cells(1, 12).Value = "Total_stock_volume"

lastrowA = ws.cells(Rows.Count, "A").End(xlUp).Row
open_Price = ws.Range("C2")

'loop to search through symbols

For i = 2 To lastrowA

'If Ticker changes then get Ticket value

If ws.cells(i, 1).Value <> ws.cells(i + 1, 1).Value Then

'print Values in column i

Ticker = ws.cells(i, 1).Value
total_stock_volume = totalstock + ws.cells(i, 7).Value

ws.Range("i" & Ticker_Row).Value = Ticker
ws.Range("L" & Ticker_Row).Value = total_stock_volume

'calculate yearly_change and Percent_change, as 0 value

yearly_change = ws.cells(i, 6) - open_Price
ws.Range("j" & Ticker_Row).Value = yearly_change


If open_Price = 0 Then
 percent_change = 0

Else

percent_change = yearly_change / open_Price

End If

ws.Range("K" & Ticker_Row).Value = percent_change

'reset variable to chane row,total volume for ticker and change the new opn_Price for ticker
Ticker_Row = Tiicker_row + 1
total_stock_volume = 0
open_Price = ws.cells(1 + i, 3)

Else

total_stock_volume = total_stock_volume + ws.cells(i, 7).Value

End If

Next i

'conditional formatting for Positive and negative values.

lr_yearly_change = ws.cells(Rows.Count, 10).End(x1up).Row

For r = 2 To lastrowA

If ws.cells(r, 10).Value < 0 Then
ws.cells(r, 10).Interior.ColorIndex = 4

ElseIf ws.cells(r, 10).Value > 0 Then
ws.cells(r, 10).Interior.ColorIndex = 3

End If

Next r


'name the greatest increase and decrease value and ticker

ws.cells(2, 15).Value = "GreatestIncrease%"
ws.cells(3, 15).Value = "GreatestDecrease%"
ws.cells(4, 15).Value = "GreatestStockVolume"
ws.cells(1, 16).Value = "ticker"
ws.cells(1, 17).Value = "value"

'define cells
Change = ws.cells("K:K")
Stockvolume = ws.cells("L:L")


'Set Variables
greatestincrease = WorksheetFunction.Max(Change)
greatestdecrease = WorksheetFunction.Min(Change)
Volume = WorksheetFunction.Max(Stockvolume)

For s = 2 To lastrowA

'determine Max and Min percent change

If ws.cells(s, 11).Value = Application.worksheetfuntion.Min(ws.cells("k2:K" & lr_percent_change)) Then


ws.cells(5, 17).Value = ws.cells(s, 11).Value
ws.cells(5, 16).Value = ws.cells(s, 9).Value

'determin max total stock volume

ElseIf ws.cells(s, 12).Value = Application.WorksheetFunction.Max(ws.Range("l2:L" & lr_percent_change)) Then

ws.cells(4, 17).Value = ws.cells(s, 12).Value
ws.cells(4, 16).Value = ws.cells(s, 9).Value

End If

Next s

Next ws

End Sub

