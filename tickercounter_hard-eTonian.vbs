Sub tickercount_hard()

Dim ws as Worksheet

For each ws in Worksheets

    Dim lastrow As Long
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percentage Change"
    ws.Range("L1").Value = "Total Stock Volume"
    Dim ticker As String
    Dim stocktotal As Double
    stocktotal = 0
    Dim summaryrow As Integer
    summaryrow = 2
    Dim opentick As Double
    opentick = 0
    Dim closetick As Double
    closetick = 0
    
    For i = 2 To lastrow

        If ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then
            opentick = ws.Cells(i, 3).Value
            stocktotal = stocktotal + ws.Cells(i, 7).Value
        ElseIf ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            closetick = ws.Cells(i, 6).Value
            ticker = ws.Cells(i, 1).Value
            stocktotal = stocktotal + ws.Cells(i, 7).Value
            ws.Range("I" & summaryrow).Value = ticker
            ws.Range("L" & summaryrow).Value = stocktotal
            ws.Range("J" & summaryrow).Value = (closetick - opentick)
            If ws.Range("J" & summaryrow).Value > 0 Then
                ws.Range("J" & summaryrow).Interior.ColorIndex = 4
            Else
                ws.Range("J" & summaryrow).Interior.ColorIndex = 3
            End If
            If opentick = 0 Then
                ws.Range("K" & summaryrow).Value = 1
            Else
                ws.Range("K" & summaryrow).Value = (closetick - opentick) / opentick
            End If
            ws.Range("K" & summaryrow).NumberFormat = "0.00%"
            summaryrow = summaryrow + 1
            stocktotal = 0
        Else
            stocktotal = stocktotal + ws.Cells(i, 7).Value
        End If

    Next i

    lastrow = ws.Cells(Rows.Count, 9).End(xlUp).Row
    ws.Range("N2").Value = "Greatest % Increase"
    ws.Range("N3").Value = "Greatest % Decrease"
    ws.Range("N4").Value = "Greatest Total Vol."
    ws.Range("O1").Value = "Ticker"
    ws.Range("P1").Value = "Value"
    dim max as Double
    max = 0
    
    For j = 2 to lastrow

        If ws.Cells(j, 11).Value > max Then
            max = ws.Cells(j, 11).Value
            ws.Range("P2").Value = max
            ws.Range("P2").NumberFormat = "0.00%"
            ticker = ws.Cells(j, 9).Value
            ws.Range("O2").Value = ticker
        End if
        
    Next J
    
    max = 0
    
    For k = 2 to lastrow
    
        If ws.Cells(k, 11).Value < max Then
            max = ws.Cells(k, 11).Value
            ws.Range("P3").Value = max
            ws.Range("P3").NumberFormat = "0.00%"
            ticker = ws.Cells(k, 9).Value
            ws.Range("O3").Value = ticker
        End If

    Next K
    
    max = 0
    
    For l = 2 to lastrow

        If ws.Cells(l, 12).Value > max Then
            max = ws.Cells(l, 12).Value
            ws.Range("P4").Value = max
            ticker = ws.Cells(l, 9).Value
            ws.Range("O4").Value = ticker
        End if

    Next l

ws.Cells.Columns.AutoFit

summaryrow = 2

Next ws

End Sub