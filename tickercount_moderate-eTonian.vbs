Sub tickercount_moderate()

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

    Cells.Columns.AutoFit
    
Next ws

End Sub

