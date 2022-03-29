Sub stocksAnalyze()
    
    '' sets up column headers
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
    
    '' ticker name
    Dim ticker As String
    ticker = Cells(2, 1).Value
    
    '' which ticker we're working on, start at 2
    Dim summaryRow As Integer
    summaryRow = 2
    
    '' year open value for tickers
    Dim startValue As Double
    yearStart = Cells(2, 3).Value
    
    '' volumeSum per ticker
    Dim volumeSum As Double
    volumeSum = 0
    
    '' determines last row on sheet
    rowMax = Cells(Rows.Count, 1).End(xlUp).Row

    For i = 2 To rowMax
        ' on a ticker's last row
        If Cells(i + 1, 1).Value <> Cells(i, 1) Then
            ' record ticker name
            Cells(summaryRow, 9) = Cells(i, 1)
            ' record yearly change and format
            Cells(summaryRow, 10) = Cells(i, 6) - yearStart
            If ((Cells(i, 6) / yearStart) - 1) * 100 < 0 Then
                Cells(summaryRow, 10).Interior.ColorIndex = 3
                Else
                Cells(summaryRow, 10).Interior.ColorIndex = 4
            End If
            ' record yearly change percent
           Cells(summaryRow, 11) = Round(((Cells(i, 6) / yearStart) - 1) * 100, 2)
            ' record total volume
           Cells(summaryRow, 12) = volumeSum + Cells(i, 7)
            ' update variables for next ticker
            summaryRow = summaryRow + 1
            yearStart = Cells(i + 1, 6)
            volumeSum = 0
            
            Else
            ' add to ticker's volume counter
            volumeSum = volumeSum + Cells(i, 7)
            
        End If

    Next i

End Sub
