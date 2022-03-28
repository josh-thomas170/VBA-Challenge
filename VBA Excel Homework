Sub MultipleYearStockData()

For Each ws In Worksheets


Dim WorksheetName As String
Dim i As Long
Dim j As Long
Dim EndRowA As Long
Dim TickerCounter As Long


WorksheetName = ws.Name
j = 2
TickerCounter = 2
EndRowA = ws.Cells(Rows.Count, 1).End(xlUp).Row


ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Percentage Change"
ws.Cells(1, 12).Value = "Total Stock Volume"


For i = 2 To EndRowA


If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

ws.Cells(TickerCounter, 9).Value = ws.Cells(i, 1).Value
ws.Cells(TickerCounter, 10).Value = ws.Cells(i, 6).Value - ws.Cells(j, 3).Value

If ws.Cells(TickerCounter, 10).Value < 0 Then

ws.Cells(TickerCounter, 10).Interior.ColorIndex = 3

Else: ws.Cells(TickerCounter, 10).Interior.ColorIndex = 50

End If


If ws.Cells(j, 3).Value <> 0 Then

PercentageChange = ((ws.Cells(i, 6).Value - ws.Cells(j, 3).Value) / ws.Cells(j, 3).Value)
ws.Cells(TickerCounter, 11).Value = Format(PercentageChange, "Percent")
                    
Else: ws.Cells(TickerCounter, 11).Value = Format(0, "Percent")
                    
End If
                    

ws.Cells(TickerCounter, 12).Value = WorksheetFunction.Sum(Range(ws.Cells(j, 7), ws.Cells(i, 7)))

TickerCounter = TickerCounter + 1
                
j = i + 1
                
End If


Next i


Next ws


End Sub



