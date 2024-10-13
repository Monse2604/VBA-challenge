# VBA-challenge
Sub AnalyzeStockData()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim ticker As String
    Dim openPrice As Double, closePrice As Double
    Dim vol As Double
    Dim quarterlyChange As Double
    Dim percentageChange As Double
    Dim totalVolume As Double
    Dim results As Collection
    Dim bestIncrease As Double, bestDecrease As Double, bestVolume As Double
    Dim bestIncreaseTicker As String, bestDecreaseTicker As String, bestVolumeTicker As String

    Set results = New Collection

    For Each ws In ThisWorkbook.Worksheets
        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        
        For i = 2 To lastRow
            ticker = ws.Cells(i, 1).Value
            openPrice = ws.Cells(i, 3).Value
            closePrice = ws.Cells(i, 6).Value
            vol = ws.Cells(i, 7).Value
            
            ' Calculate quarterly change and percentage change
            quarterlyChange = closePrice - openPrice
            If openPrice <> 0 Then
                percentageChange = (quarterlyChange / openPrice) * 100
            Else
                percentageChange = 0
            End If
            
            ' Accumulate total volume
            totalVolume = totalVolume + vol
            
            ' Store results for the current stock
            results.Add Array(ticker, quarterlyChange, percentageChange, vol)
        Next i

        ' Find the stock with the greatest increase, decrease, and total volume
        For i = 1 To results.Count
            If results(i)(2) > bestIncrease Then
                bestIncrease = results(i)(2)
                bestIncreaseTicker = results(i)(0)
            End If
            If results(i)(2) < bestDecrease Then
                bestDecrease = results(i)(2)
                bestDecreaseTicker = results(i)(0)
            End If
            bestVolume = Application.Max(bestVolume, results(i)(3))
            If bestVolume = results(i)(3) Then
                bestVolumeTicker = results(i)(0)
            End If
        Next i

        ' Output the results for the sheet
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Quarterly Change"
        ws.Cells(1, 11).Value = "Percentage Change"
        ws.Cells(1, 12).Value = "Total Volume"
        
        For i = 1 To results.Count
            ws.Cells(i + 1, 9).Value = results(i)(0)
            ws.Cells(i + 1, 10).Value = results(i)(1)
            ws.Cells(i + 1, 11).Value = results(i)(2)
            ws.Cells(i + 1, 12).Value = results(i)(3)
        Next i

        ' Format for colors
        With ws.Range(ws.Cells(2, 10), ws.Cells(results.Count + 1, 10))
            .NumberFormat = "0.00%"
            For i = 2 To results.Count + 1
                If ws.Cells(i, 10).Value > 0 Then
                    ws.Cells(i, 10).Interior.Color = RGB(144, 238, 144) ' Light Green for positive
                ElseIf ws.Cells(i, 10).Value < 0 Then
                    ws.Cells(i, 10).Interior.Color = RGB(255, 99, 71) ' Light Red for negative
                Else
                    ws.Cells(i, 10).Interior.ColorIndex = xlNone ' No color for zero
                End If
            Next i
        End With
        
        ' best stocks
        ws.Cells(1, 14).Value = "Greatest % Increase: " & bestIncreaseTicker & " (" & bestIncrease & "%)"
        ws.Cells(2, 14).Value = "Greatest % Decrease: " & bestDecreaseTicker & " (" & bestDecrease & "%)"
        ws.Cells(3, 14).Value = "Greatest Total Volume: " & bestVolumeTicker & " (" & bestVolume & ")"

        ' Clear the collection for the next sheet
        Set results = New Collection
        totalVolume = 0
    Next ws
End Sub
