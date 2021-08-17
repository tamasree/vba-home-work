Attribute VB_Name = "Max_Increase_and_Decrease"
Sub StockAnalysis3()

For Each ws In Worksheets

Dim MaxIncrease As Double
Dim MaxDecrease As Double
Dim lrow As Long
Dim Myrange As Range
Dim cell As Range

' Find lastrow of summarytable'

lrow = ws.Range("L" & Rows.Count).End(xlUp).Row

'set range to find maxincrease'

Set Myrange = ws.Range("L2:L" & lrow)

'Find max value in the range'

MaxIncrease = ws.Application.WorksheetFunction.Max(Myrange)

MaxDecrease = ws.Application.WorksheetFunction.Min(Myrange)

'Print maxincrease and max decrease'

ws.Range("R2").Value = MaxIncrease
ws.Range("R2").NumberFormat = "0.0%"

ws.Range("R3").Value = MaxDecrease
ws.Range("R3").NumberFormat = "0.0%"

' Print Ticker for maxincrease and max decrease'
For Each cell In Myrange
  
    If cell.Value = MaxIncrease Then
    
    ws.Range("Q2").Value = cell.Offset(, -2).Value
     
    ElseIf cell.Value = MaxDecrease Then
    
    ws.Range("Q3").Value = cell.Offset(, -2).Value
     
    End If

Next cell

Next ws
End Sub


'

