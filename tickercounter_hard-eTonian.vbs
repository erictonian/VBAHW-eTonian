Sub tickercount_hard()

Dim ws As Worksheet

For Each ws In Worksheets

    Dim lastrow As Long
    lastrow = Cells(Rows.Count, 1).End(xlUp).Row
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percentage Change"
    Range("L1").Value = "Total Stock Volume"
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

        If Cells(i - 1, 1).Value <> Cells(i, 1).Value Then
            opentick = Cells(i, 3).Value
            stocktotal = stocktotal + Cells(i, 7).Value
        ElseIf Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            closetick = Cells(i, 6).Value
            ticker = Cells(i, 1).Value
            stocktotal = stocktotal + Cells(i, 7).Value
            Range("I" & summaryrow).Value = ticker
            Range("L" & summaryrow).Value = stocktotal
            Range("J" & summaryrow).Value = (closetick - opentick)
            If Range("J" & summaryrow).Value > 0 Then
                Range("J" & summaryrow).Interior.ColorIndex = 4
            Else
                Range("J" & summaryrow).Interior.ColorIndex = 3
            End If
            If opentick = 0 Then
                Range("K" & summaryrow).Value = 1
            Else
                Range("K" & summaryrow).Value = (closetick - opentick) / opentick
            End If
            Range("K" & summaryrow).NumberFormat = "0.00%"
            summaryrow = summaryrow + 1
            stocktotal = 0
        Else
            stocktotal = stocktotal + Cells(i, 7).Value
        End If

    Next i

    lastrow = Cells(Rows.Count, 9).End(xlUp).Row
    Range("N2").Value = "Greatest % Increase"
    Range("N3").Value = "Greatest % Decrease"
    Range("N4").Value = "Greatest Total Vol."
    Range("O1").Value = "Ticker"
    Range("P1").Value = "Value"
    dim max as Double
    max = 0
    
    For j = 2 to lastrow
        If Cells(j, 11).Value > max Then
            max = Cells(j, 11).Value
            Range("P2").Value = max
            Range("P2").NumberFormat = "0.00%"
            ticker = Cells(j, 9).Value
            Range("O2").Value = ticker
        End if
    Next J
    
    max = 0
    
    For k = 2 to lastrow
    
        If Cells(k, 11).Value < max Then
            max = Cells(k, 11).Value
            Range("P3").Value = max
            Range("P3").NumberFormat = "0.00%"
            ticker = Cells(k, 9).Value
            Range("O3").Value = ticker
        End If
    Next K
    
    max = 0
    
    For l = 2 to lastrow
        If Cells(l, 12).Value > max Then
            max = Cells(l, 12).Value
            Range("P4").Value = max
            ticker = Cells(l, 9).Value
            Range("O4").Value = ticker
        End if
    Next l

Cells.Columns.AutoFit

Next ws

End Sub