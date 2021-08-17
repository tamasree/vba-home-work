Attribute VB_Name = "Summarytable"
Sub StockAnalisys1()

For Each ws In Worksheets

'Adding Hearders for Summarytable'

ws.Range("J1").Value = "Ticker"
ws.Range("K1").Value = "Yearly Change"
ws.Range("L1").Value = "Percent Change"
ws.Range("M1").Value = "Total Stock Volume"

'Setting summary table row count as 2 after header'

Summarytable_Row = 2

'Find last row for the data'

Dim lastrow As Long
lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

Dim Total_Volume  As Long

'Setting Total volume as 0 for first stock of the sheet'

Total_Volume = 0

Dim OpenValue As Double
Dim CloseValue As Double
Dim Percentchange As Double

OpenValue = ws.Range("C2").Value

For i = 2 To lastrow

     If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
     
'Determine Total value of a stock at the end of a year'

             ws.Range("J" & Summarytable_Row).Value = ws.Cells(i, 1).Value

             Total = Total + ws.Cells(i, 7).Value
             
            'Print total value in summary table'
            
             ws.Range("M" & Summarytable_Row).Value = Total
             
             'Determine Close value of that stock'
             
            CloseValue = ws.Cells(i, 6).Value
            
            'Find yearly change of the stock based on open and closed value'
            
            
            YearlyChange = CloseValue - OpenValue
            
            'Print yearlychange in summary table
            
            ws.Range("K" & Summarytable_Row).Value = YearlyChange
      
         'Determine percentage change on open value of a stock at the end of a year'
       
            If OpenValue <> 0 Then

                    Percentchange = YearlyChange / OpenValue
            
                      'Print percentage change in summary table in % format'
            
                      ws.Range("L" & Summarytable_Row).Value = Percentchange
                      ws.Range("L" & Summarytable_Row).NumberFormat = "0.0%"
            
            Else
                        ws.Range("L" & Summarytable_Row).Value = " "
                        
            End If
            
            'Conditional Formatting Based on +ve and -ve percentage'
            
            If Percentchange < 0 Then
            ws.Range("L" & Summarytable_Row).Interior.Color = RGB(255, 0, 0)
            
            ElseIf Percentchange >= 0 Then
            
            ws.Range("L" & Summarytable_Row).Interior.Color = RGB(0, 255, 0)
            
             End If
             
            'Increase summary table count to put the next stock in next row of summary table'
            
            Summarytable_Row = Summarytable_Row + 1
            
            'Reset total volume to 0 for the next stock'

             Total = 0
             
             'Reset open value for next stock'
             
             OpenValue = ws.Cells(i + 1, 3).Value
             
   Else
             Total = Total + ws.Cells(i, 7).Value
             
  End If

Next i

Next ws

MsgBox "SummaryTable Done For All Sheets"

End Sub

