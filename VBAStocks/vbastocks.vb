Sub tickerTool():
    'Find last row of sheet
    Dim lastRow As Double
    lastRow = ActiveSheet.Cells(Rows.Count, "B").End(xlUp).Row
    'Test last row output
    'MsgBox (lastRow)
    'Make array to store Ticker, Yearly Change, Percent Change, Total Stock Volume
    Dim tickerArray() As String
    Dim yearlyChange() As Double
    Dim percentChange() As Double
    Dim totalStockVolume() As Double
    Dim x As Double

    For i = 2 To lastRow
        Dim priceChange As Double
        Dim totalVolume As Double
        Dim numOfRowsToLast As Double
        Dim yearPercentChange As Double

        'If next ticker matches previous ticker
        If (Cells(i, 1).Value = Cells(i + 1, 1)) Then
            'Add volumes with the same ticker
            totalVolume = Cells(i, 7) + totalVolume
            'Record rows until last in ticker
            numOfRowsToLast = numOfRowsToLast + 1

        'If next ticker doesn't match previous ticker
        ElseIf (Cells(i, 1).Value <> Cells(i + 1)) Then
            priceChange = Cells(i, 6) - Cells(i - numOfRowsToLast, 3)
            totalVolume = Cells(i, 7) + totalVolume
            '[(Old Price - New Price)/Old Price]
            If Cells(i - numOfRowsToLast, 3) > 0 Then
                yearPercentChange = -((Cells(i - numOfRowsToLast, 3) - Cells(i, 6)) / Cells(i - numOfRowsToLast, 3))
            Else
                yearPercentChange = 0
            End If

            'Store ticker names in ticker array
            ReDim Preserve tickerArray(x)
            tickerArray(x) = Cells(i, 1).Value

            'Store totalChange in yearlyChange array
            ReDim Preserve yearlyChange(x)
            yearlyChange(x) = priceChange

            'Store percentage change in percentChange array
            ReDim Preserve percentChange(x)
            percentChange(x) = yearPercentChange

            'Store total volume in totalVolume array
            ReDim Preserve totalStockVolume(x)
            totalStockVolume(x) = totalVolume

            'Add 1 to current array index
            x = x + 1


            'zero out values
            'totalChange = 0
            priceChange = 0
            totalVolume = 0
            initialVolume = 0
            numOfRowsToLast = 0
        End If

    Next i
    'MsgBox tickerArray(270)
    'MsgBox Format(yearlyChange(0), "#,##0.00")
    'MsgBox totalStockVolume(0)
    'MsgBox y

    'Add columns with specific headers
        Range("I1").Value = "Ticker"
        Range("J1").Value = "Yearly Change"
        Range("K1").Value = "Percent Change"
        Range("L1").Value = "Total Stock Volume"


    For a = 2 To x

        Cells(a, 9).Value = tickerArray(a - 2)
        Cells(a, 10).Value = Format(yearlyChange(a - 2), "#,##0.00")

        'Check if Yearly Change is less than 0
        If Cells(a, 10).Value < 0 Then
            'Interior color is red if negative
            Cells(a, 10).Interior.ColorIndex = 3
        Else
            'Interior color is green if positive
            Cells(a, 10).Interior.ColorIndex = 4
        End If
        Cells(a, 11).Value = FormatPercent(percentChange(a - 2))
        Cells(a, 12).Value = totalStockVolume(a - 2)
    Next a

    'Challenge: Find greatest % increase, decrease and greatest total volume
    Dim greatestIncrease As Double
    Dim greatestDecrease As Double
    Dim greatestTotalVolume As Double
    Dim maxIndex As Long
    Dim minIndex As Long
    Dim volMaxIndex As Long
        'Find max from percentChange array
        greatestIncrease = WorksheetFunction.Max(percentChange())
        'Find index at max
        maxIndex = Application.Match(Application.Max(percentChange()), percentChange, 0) - 1
        'Find min from percentChange array
        greatestDecrease = WorksheetFunction.Min(percentChange())
        'Find index at min
        minIndex = Application.Match(Application.Min(percentChange()), percentChange, 0) - 1
        'Find max in totalStockVolume array
        greatestTotalVolume = WorksheetFunction.Max(totalStockVolume())
        'Find index at max
        volMaxIndex = Application.Match(Application.Max(totalStockVolume()), totalStockVolume, 0) - 1
        'MsgBox greatestIncrease
        'Set row & column labels
        Range("P1").Value = "Ticker"
        Range("Q1").Value = "Value"
        Range("O2").Value = "Greatest % Increase"
        Range("O3").Value = "Greatest % Decrease"
        Range("O4").Value = "Greatest Total Volume"
        Range("P2").Value = tickerArray(maxIndex)
        Range("Q2").Value = FormatPercent(greatestIncrease)
        Range("P3").Value = tickerArray(minIndex)
        Range("Q3").Value = FormatPercent(greatestDecrease)
        Range("P4").Value = tickerArray(volMaxIndex)
        Range("Q4").Value = greatestTotalVolume
        

'       Some debugging stuff...
'        Dim p As Long
'        Debug.Print WorksheetFunction.Max(percentChange())
'        p = Application.Match(Application.Max(percentChange()), percentChange, 0) - 1
'        Debug.Print p
'        Debug.Print tickerArray(p)
        'Debug.Print WorksheetFunction.Min(percentChange())

End Sub