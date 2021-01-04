Sub stock_market_report()
Dim i, a, b As Integer
Dim openyear, closeyear, perc_change As Double
Dim year_change As Double
Dim yearstock As String
Dim Ticker As String
Dim Volumen As Double
Dim maxperc, minperc, maxvolumen As Double

For Each ws In Worksheets
 'initial Variables
 a = 2
 openyear = ws.Cells(2, 3).Value
 closeyear = 0
 Volumen = 0
 lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
 'Create top labels
 ws.Range("J1").Value = "Ticker"
 ws.Range("K1").Value = "Year Change"
 ws.Range("L1").Value = "Percent Change"
 ws.Range("M1").Value = "Total Stock Volumen"
 
'for for each row on the ws
 For i = 2 To lastrow
    Ticker = ws.Cells(i, 1).Value
    If Ticker = ws.Cells(i + 1, 1).Value Then
        Volumen = ws.Cells(i, 7).Value + Volumen
    Else
        Volumen = ws.Cells(i, 7).Value + Volumen
        closeyear = ws.Cells(i, 6).Value
        year_change = closeyear - openyear
        If (year_change < 0) Then
            ws.Cells(a, 11).Interior.ColorIndex = 3
        Else
            ws.Cells(a, 11).Interior.ColorIndex = 4
        End If
        If (openyear <> 0) Then
             perc_change = (year_change / openyear)
        End If
        ws.Cells(a, 10).Value = Ticker
        ws.Cells(a, 11).Value = year_change
        ws.Cells(a, 12).Value = perc_change
        ws.Cells(a, 12).NumberFormat = "0.00%"
        ws.Cells(a, 13).Value = Volumen
        a = a + 1
        closeyear = 0
        year_change = 0
        perc_change = 0
        Volumen = 0
        openyear = ws.Cells(i + 1, 3).Value
    End If
 Next i
 'get the greater end smaller values of the new values
 myrange1 = "L2:L" + VBA.Format(ws.Cells(Rows.Count, 10).End(xlUp).Row)
 myrange2 = "M2:M" + VBA.Format(ws.Cells(Rows.Count, 11).End(xlUp).Row)
 ws.Cells(1, 16).Value = "Greater % Increase"
 ws.Cells(2, 16).Value = "Greater % Decrease"
 ws.Cells(1, 16).Interior.ColorIndex = 3
 ws.Cells(2, 16).Interior.ColorIndex = 4
 ws.Cells(3, 16).Value = "Greater Total volumen"
 maxperc = WorksheetFunction.Max(ws.Range(myrange1))
 ws.Cells(1, 17).Value = maxperc
 ws.Cells(1.17).NumberFormat = "0.00%"
 minperc = WorksheetFunction.Min(ws.Range(myrange1))
 ws.Cells(2, 17).Value = minperc
 ws.Cells(2, 17).NumberFormat = "0.00%"
 maxvolumen = WorksheetFunction.Max(ws.Range(myrange2))
 ws.Cells(3, 17).Value = maxvolumen
 ws.Columns("J:Q").AutoFit
Next ws

End Sub

Sub Clear()
Dim i As Integer

For Each ws In Worksheets
 lastrow = ws.Cells(Rows.Count, 10).End(xlUp).Row
 For i = 1 To lastrow
    ws.Cells(i, 10).Value = ""
    ws.Cells(i, 11).Value = ""
    ws.Cells(i, 12).Value = ""
    ws.Cells(i, 13).Value = ""
    ws.Cells(i, 11).Interior.ColorIndex = 0
 Next i
  ws.Cells(1, 17).Value = ""
  ws.Cells(2, 17).Value = ""
  ws.Cells(3, 17).Value = ""
  ws.Cells(1, 16).Value = ""
  ws.Cells(2, 16).Value = ""
  ws.Cells(3, 16).Value = ""
  ws.Cells(1, 16).Interior.ColorIndex = 0
  ws.Cells(2, 16).Interior.ColorIndex = 0
  ws.Columns("J:Q").AutoFit
Next ws
End Sub
