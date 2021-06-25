Sub MarketCleanup()
    Dim tickerSymbol As String
    Dim nextTickerSymbol As String
    Dim lastRow As Long
    Dim i As Long
    Dim totalStockVolume As Variant
    Dim openingPrice As Double
    Dim closingPrice As Double
    Dim yearlyChange As Double
    Dim percentageChange As Double
    Dim greatestIncreaseValue As Double
    Dim greatestDecreaseValue As Double
    Dim greatestTotalVolumeValue As Variant
    Dim greatestIncreaseTicker As String
    Dim greatestDecreaseTicker As String
    Dim greatestTotalVolumeTicker As String
    
    
    
    
    For Each ws In Worksheets
        'Initialize the values for the first row of the worksheet
        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        tickerSymbol = ws.Cells(2, 1).Value
        openingPrice = ws.Cells(2, 3).Value
        nextTickerRow = 2
        totalStockVolume = 0
        percentageChange = 0
        
        'Initialize Greatest summary
        greatestIncreaseValue = 0
        greatestDecreaseValue = 2147483647
        greatestTotalVolumeValue = 0
        greatestIncreaseTicker = ""
        greatestDecreaseTicker = ""
        greatestTotalVolumeTicker = ""
        
        
        For i = 2 To lastRow
            totalStockVolume = totalStockVolume + ws.Cells(i, 7).Value
            
            If ws.Cells(i + 1, 1).Value <> tickerSymbol Then
                yearlyChange = ws.Cells(i, 6).Value - openingPrice
                
                ws.Cells(nextTickerRow, 9).Value = tickerSymbol
                ws.Cells(nextTickerRow, 10).Value = yearlyChange
                
                'YearlyChange conditional format
                If (yearlyChange > 0) Then
                     ws.Cells(nextTickerRow, 10).Interior.ColorIndex = 4
                Else
                    ws.Cells(nextTickerRow, 10).Interior.ColorIndex = 3
                End If
                
                'PercentageChanged
                If (openingPrice > 0) Then
                    percentageChange = yearlyChange / openingPrice
                    ws.Cells(nextTickerRow, 11).Value = percentageChange
                Else
                    ws.Cells(nextTickerRow, 11).Value = "N/A"
                End If
                
                'Total Stock Volume
                ws.Cells(nextTickerRow, 12).Value = totalStockVolume
                
                'Greatest Increase
                If (percentageChange > greatestIncreaseValue) Then
                    greatestIncreaseValue = percentageChange
                    greatestIncreaseTicker = tickerSymbol
                End If
                                
                'Greatest Decrease
                If (percentageChange < greatestDecreaseValue) Then
                    greatestDecreaseValue = percentageChange
                    greatestDecreaseTicker = tickerSymbol
                End If
                
                'Greatest Total Volume
                If (totalStockVolume > greatestTotalVolumeValue) Then
                    greatestTotalVolumeValue = totalStockVolume
                    greatestTotalVolumeTicker = tickerSymbol
                End If
                
                
                
                'Initialize the values for the next ticker
                totalStockVolume = 0
                tickerSymbol = ws.Cells(i + 1, 1).Value
                openingPrice = ws.Cells(i + 1, 3).Value
                nextTickerRow = nextTickerRow + 1
            End If
        Next i
        ws.Cells(2, 16).Value = greatestIncreaseTicker
        ws.Cells(2, 17).Value = greatestIncreaseValue
        ws.Cells(3, 16).Value = greatestDecreaseTicker
        ws.Cells(3, 17).Value = greatestDecreaseValue
        ws.Cells(4, 16).Value = greatestTotalVolumeTicker
        ws.Cells(4, 17).Value = greatestTotalVolumeValue
    Next ws
    
    MsgBox ("Done")
    
End Sub
