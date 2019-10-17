Sub stocks()

    Dim ticker, pTicker As String
    
    Dim firstOpening, finalClosing, tVolumeSubTotal As Double
    tVolumeSubTotal = 0

    Dim tickerCount, lRow As Long
    
    ' looping through all the worksheets
    For w = 1 to Worksheets.Count
        Worksheets(w).Activate

        'get the last row with data in each worksheet
        lRow = Cells.Find(What:="*", _
                    After:=Range("A1"), _
                    LookAt:=xlPart, _
                    LookIn:=xlFormulas, _
                    SearchOrder:=xlByRows, _
                    SearchDirection:=xlPrevious, _
                    MatchCase:=False).Row

        tickerCount = 1

        'will enter data in these columns
        Cells(1, 9).Value2 = "Ticker"
        Range("I:I").NumberFormat = "@"
        Cells(1, 10).Value2 = "Yearly Change"
        'the homework screenshots had more significant digits, I don't know where those came from. there are
        'two digits. 
        Range("J:J").NumberFormat = "0.00"
        Cells(1, 11).Value2 = "Percent Change"
        Range("K:K").NumberFormat = "0.00%"
        Cells(1, 12).Value2 = "Total Stock Volume"
        Range("L:L").NumberFormat = "0"
        
        'first row only
        ticker = Cells(2, 1).Value2
        'enter name for the first ticker
        Cells(2, 9).Value2 = ticker
        pTicker = ticker

        firstOpening = Cells(2, 3).Value2
        
        ' this logic assumes a sorted list of tickers. If the list were unsorted, I'd probably create an array
        ' and search through that each time to check if each ticker had already been added to it. Would also need
        ' to store subtotals for each ticker without overwriting them once a new ticker is found, and to find
        ' earliest and latest dates.
        For i = 2 To lRow + 1
            ticker = Cells(i, 1).Value2
            If ticker <> pTicker Then
                tickerCount = tickerCount + 1

                'get the final closing price for the previous row
                finalClosing = Cells(i - 1, 6).Value2

                'enter the name for the next ticker. This will enter an extra empty string after the last row
                Cells(tickerCount + 1, 9).Value2 = ticker

                'Michael pointed out that some of the stocks opened at 0, causing divide-by-zero errors,
                'so we'll check for that first
                If firstOpening <> 0 Then
                    Cells(tickerCount, 10) = finalClosing - firstOpening
                    Cells(tickerCount, 11) = (finalClosing / firstOpening)-1
                    Cells(tickerCount, 12) = tVolumeSubTotal
                Else
                    Cells(tickerCount, 10) = finalClosing - firstOpening
                    Cells(tickerCount, 11) = 0
                    Cells(tickerCount, 12) = tVolumeSubTotal
                End If

                'new first opening price for next ticker
                firstOpening = Cells(i, 3).Value2

                'rest volume subtotal for next ticker
                tVolumeSubTotal = 0
            End If
            tVolumeSubTotal = tVolumeSubTotal + Cells(i, 7).Value2
            pTicker = ticker
        Next i

        'I hate this syntax, but it's more compact than looping through the whole column
        Range("J:J").FormatConditions.Delete
        Range("J1:J"&tickerCount).FormatConditions.Add(xlCellValue, xlGreaterEqual, "=0").Interior.ColorIndex = 4
        Range("J1:J"&tickerCount).FormatConditions.Add(xlCellValue, xlLess, "=0").Interior.ColorIndex = 3
    
        'Challenge
        Cells(2,14).Value2 = "Greatest % Increase"
        Cells(3,14).Value2 = "Greatest % Decrease"
        Cells(4,14).Value2 = "Greatest Total Volume"
        Cells(1,15).Value2 = "Ticker"
        Range("O1:O4").NumberFormat = "@"
        Cells(1,16).Value2 = "Value"
        Range("P2:P3").NumberFormat = "0.00%"
        Range("P4").NumberFormat = "0"

        Dim maxIncrease, maxDecrease, maxVolume As Double
        maxIncrease = 0
        maxDecrease = 0
        maxVolume = 0

        Dim tickerMaxIncrease, tickerMaxDecrease, tickerMaxVolume As String
        
        For i = 2 to tickerCount
            If Cells(i,11).Value2 > maxIncrease Then
                maxIncrease = Cells(i,11).Value2
                tickerMaxIncrease = Cells(i,9).Value2
            ElseIf Cells(i,11).Value2 < maxDecrease Then
                maxDecrease = Cells(i,11).Value2
                tickerMaxDecrease = Cells(i,9).Value2
            End If
            If Cells(i,12).Value2 > maxVolume Then
                maxVolume = Cells(i,12).Value2
                tickerMaxVolume = Cells(i,9).Value2
            End If
        Next i

        Cells(2,15) = tickerMaxIncrease
        Cells(2,16) = maxIncrease
        Cells(3,15) = tickerMaxDecrease
        Cells(3,16) = maxDecrease
        Cells(4,15) = tickerMaxVolume
        Cells(4,16) = maxVolume

    Next w
End Sub