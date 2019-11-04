Attribute VB_Name = "Module1"
Sub vbastocks()

'CreateVariables
Dim ws As Worksheet
Dim ticker As String
Dim yearly_change As Double
Dim price_open As Double
Dim price_close As Double
Dim percent_change As Double
Dim previous_amount As Long
Dim total_volume As Double
Dim lastrow_ticker As Long
Dim lastrow_percent_change As Long
Dim lastrow As Long
Dim summary_table_row As Long
Dim i As Long

'Loop through each ws & add/label headers
For Each ws In Worksheets
ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "Total Stock Volume"
ws.Cells(2, 14).Value = "Greatest % Increase"
ws.Cells(3, 14).Value = "Greatest % Decrease"
ws.Cells(4, 14).Value = "Greatest Total Volume"
ws.Cells(1, 15).Value = "Ticker"
ws.Cells(1, 16).Value = "Value"

'Determine last row in ticker column (A)
lastrow_ticker = ws.Cells(Rows.Count, "A").End(xlUp).Row

'Keep track of location for each variable in summary table & set initial variables to 0
summary_table_row = 2
previous_amount = 2
total_volume = 0
price_open = 0
price_close = 0
yearly_change = 0
percent_change = 0

'Loop through rows
For i = 2 To lastrow_ticker
total_volume = total_volume + ws.Cells(i, 7).Value

'If next row is same as previous..
If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
ticker = ws.Cells(i, 1).Value
price_open = ws.Range("C" & previous_amount)
price_close = ws.Range("F" & i)
yearly_change = Round(price_close - price_open, 3)

'If price open is 0 then percent change is 0. Else calculate percent change
If price_open = 0 Then
percent_change = 0
Else
percent_change = (yearly_change / price_open)
End If

'Make K column into percent format & J column into dollar format
ws.Columns("K").NumberFormat = "0.00%"
ws.Columns("J").NumberFormat = "$0.00"

'Print to summary table
ws.Range("I" & summary_table_row).Value = ticker
ws.Range("J" & summary_table_row).Value = yearly_change
ws.Range("K" & summary_table_row).Value = percent_change
ws.Range("L" & summary_table_row).Value = total_volume


'Conditional color formatting
If (yearly_change >= 0) Then
ws.Range("J" & summary_table_row).Interior.ColorIndex = 4
ElseIf (yearly_change <= 0) Then
ws.Range("J" & summary_table_row).Interior.ColorIndex = 3
End If

'Add 1 to summary table & set previous amount
summary_table_row = summary_table_row + 1
previous_amount = i + 1
End If
Next i


'Challenges:
'Determine last row in percent change (K)
lastrow_percent_change = ws.Cells(Rows.Count, 11).End(xlUp).Row

'Declare variables & set initial variables
Dim max_ticker As String
Dim min_ticker As String
max_ticker = ""
min_ticker = ""

Dim max_percent As Double
Dim min_percent As Double
max_percent = 0
min_percent = 0

Dim max_volume_ticker As String
Dim max_volume As Double
max_volume_ticker = ""
max_volume = 0

'Calculate summary table tickers & values through loop
For i = 2 To lastrow_percent_change

If ws.Range("K" & i).Value > ws.Range("P2").Value Then
max_percent = percent_change
max_ticker = ticker
End If

If ws.Range("K" & i).Value < ws.Range("P3").Value Then
min_percent = percent_change
min_ticker = ticker
End If

If ws.Range("L" & i).Value > ws.Range("P4").Value Then
max_volume = total_volume
max_volume_ticker = ticker
End If
Next i

'Print tickers & values to summary table
ws.Cells(2, 15).Value = max_ticker
ws.Cells(3, 15).Value = min_ticker
ws.Cells(4, 15).Value = max_volume_ticker
ws.Cells(2, 16).Value = max_percent
ws.Cells(3, 16).Value = min_percent
ws.Cells(4, 16).Value = max_volume

'Make P column into percent format
ws.Columns("P").NumberFormat = "0.00%"

Next ws
End Sub


