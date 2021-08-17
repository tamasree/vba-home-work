Attribute VB_Name = "MaxTotal"
Sub StockAnalysis2()

Dim MaxTotal As LongLong
Dim lrow As Long
Dim Myrange1 As Range
Dim cell As Range

For Each ws In Worksheets

' Printing headers'
ws.Range("Q1").Value = "Ticker"
ws.Range("R1").Value = "Value"

ws.Range("P2").Value = "Greatest Percentage Increase"
ws.Range("P3").Value = "Greatest Percentage Decrease"
ws.Range("P4").Value = "Greatest Total Volume"

' Find lastrow of summarytable'

lrow = ws.Range("J" & Rows.Count).End(xlUp).Row

'Search Greatest Total Volume'


For i = 2 To lrow

'set range to find max total'

Set Myrange1 = ws.Range("M2:M" & lrow)

'Find max value in the range'

MaxTotal = Application.WorksheetFunction.Max(Myrange1)

'Print Maxtotal'
ws.Range("P4").Offset(, 2).Value = MaxTotal

'Print ticker name'

For Each cell In Myrange1

If cell.Value = MaxTotal Then

ws.Range("P4").Offset(, 1).Value = cell.Offset(, -3).Value

End If

Next cell

Next i


Next ws
End Sub


