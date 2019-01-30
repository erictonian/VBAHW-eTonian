Sub tickercounteasy()

For Each ws In Worksheets

    Dim LastRow As Long
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Total Stock Volume"

    For i = 2 To LastRow

    Dim ticker As String
    Dim stocktotal As Double
    stocktotal = 0
    Dim summaryrow As Integer
    summaryrow = 2

        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            ticker = Cells(i, 1).Value
            stocktotal = stocktotal + Cells(i, 7).Value
            Range("I" & summaryrow).Value = ticker
            Range("J" & summaryrow).Value = stocktotal
            summaryrow = summaryrow + 1
            stocktotal = 0
        Else
            stocktotal = stocktotal + Cells(i, 7).Value
        End If

    Next i

Next ws

End Sub