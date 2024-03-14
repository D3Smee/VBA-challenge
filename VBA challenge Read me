'(START OF READ ME DO NOT USE THIS FIRST SUB) Sub multiple_year_stock_data 'DO NOT USE
    'Dim ws As Worksheet
    Dim lastRow As Double
    Dim i As Long
    Dim total_volume As Double
    total_volume = 0
    
    'Extract stock data (ticker symbol, opening price, closing price, volume)
    Dim ticker As String
    Dim openPrice As Double
    Dim closePrice As Double
    Dim volume As Double
    Dim yeardate As Double
    Dim Stock_info As Double
    Stock_info = 2
    ' Set the worksheet containing the stock data
    For Each ws In Worksheets
        'Set ws = ThisWorkbook.Sheets("2018")
        ' Find the last row with data in column A
        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        ' Loop through the rows to extract stock data
        For i = 2 To lastRow ' Assuming data starts from row 2
            'create loop to cycle through tickers and then create function to select the close value for the last ticker
            'ticker = ws.Cells(i, 1).Value ' Assuming ticker symbol is in column A
            '  openPrice = ws.Cells(i, 3).Value ' Assuming opening price is in column B
            ' closePrice = ws.Cells(i, 6).Value ' Assuming closing price is in column C
            'volume = ws.Cells(i, 7).Value ' Assuming volume is in column G
            'yeardate = ws.Cells(i, 2).Value ' DAte is in column 2
            
            If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
                ticker = Cells(i, 1).Value
                ws.Cells(Stock_info, "I") = ticker
                Stock_info = Stock_info + 1
                
                If ws.Cells(Stock_info, 1).Value = ws.Cells(i + 1, 1).Value Then
                'increase volume for current ticker
                total_volume = total_volume + ws.Cells(i, 7).Value
                ws.Cells(Stock_info, "L") = total_volume
                
            End If
            
            End If
            
            
                
                'Add to the volume total
                  
                'Print ticker symbol in ticker column
                ' Range("I" & Stock_info).Value = ticker
                'Print Volume in volume column
                'Range("j" & Stock_info).Value = total_volume
                'Reset volume total
                'volume_total = 0
                'print total f
                'ws.Cells(Stock_info, "L").Value = total_volume

                'Else
                '   volume_total = Cells(i, 7).Value
                'Stock_info = Stock_info + 1
                    
               ' Else
                'ws.Cells(Stock_info, "L") = total_volume
            'End If
        Next i
    Next ws
End Sub

Sub mutliyearstockdata()

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim totalVolume As Double
    Dim ticker As String
    Dim openPrice As Double
    Dim closePrice As Double
    Dim yearlyChange As Double
    Dim percentChange As Double
    Dim startPrice As Double
    Dim stockInfoRow As Long
    stockInfoRow = 2


    
    For Each ws In ThisWorkbook.Worksheets
    stockInfoRow = 2
        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        totalVolume = 0
        ' Assume first row has the initial opening price
        If lastRow >= 2 Then
            ticker = ws.Cells(2, 1).Value
            startPrice = ws.Cells(2, 3).Value ' Assuming opening price is the first record of the year
        End If
        
        For i = 2 To lastRow
            totalVolume = totalVolume + ws.Cells(i, 7).Value ' Accumulate volume
            
            If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
                ' New ticker or end of data, calculate and reset
                ticker = ws.Cells(i, 1).Value
                closePrice = ws.Cells(i, 6).Value ' Last record of the year for the ticker
                yearlyChange = closePrice - startPrice
                If startPrice <> 0 Then
                    percentChange = (yearlyChange / startPrice) * 100
                Else
                    percentChange = 0
                End If
                
                ' Output to columns I to M: Ticker, Yearly Change, Percent Change, Total Volume
                With ws
                    .Cells(stockInfoRow, "I") = ticker
                    .Cells(stockInfoRow, "J") = yearlyChange
                    .Cells(stockInfoRow, "K") = percentChange & "%"
                    .Cells(stockInfoRow, "L") = totalVolume
                End With
                
                stockInfoRow = stockInfoRow + 1
                totalVolume = 0
                
                If i + 1 <= lastRow Then
                    startPrice = ws.Cells(i + 1, 3).Value ' Set new start price for next ticker
                End If
            End If
         Next i
         
         
        Next ws
      End Sub
      
Sub GreatesMetrics()
        Dim ws As Worksheet
        Dim lastRow As Long
        Dim i As Long
        Dim totalVolume As Double
        Dim ticker As String
        Dim openPrice As Double
        Dim closePrice As Double
        Dim yearlyChange As Double
        Dim percentChange As Double
        Dim startPrice As Double
        Dim stockInfoRow As Long
    
        ' Variables for tracking the greatest metrics
        Dim greatestIncrease As Double
        Dim greatestDecrease As Double
        Dim greatestVolume As Double
        Dim tickerGreatestIncrease As String
        Dim tickerGreatestDecrease As String
        Dim tickerGreatestVolume As String

            For Each ws In ThisWorkbook.Worksheets
                ' Reset metrics for each worksheet
                stockInfoRow = 2
                greatestIncrease = 0
                greatestDecrease = 0
                greatestVolume = 0
                tickerGreatestIncrease = ""
                tickerGreatestDecrease = ""
                tickerGreatestVolume = ""
        
                lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        
                If lastRow >= 2 Then
                    ticker = ws.Cells(2, 1).Value
                    startPrice = ws.Cells(2, 3).Value
                End If
        
                For i = 2 To lastRow
                    totalVolume = totalVolume + ws.Cells(i, 7).Value
        
                    If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
                        ticker = ws.Cells(i, 1).Value
                        closePrice = ws.Cells(i, 6).Value
                        yearlyChange = closePrice - startPrice
                        If startPrice <> 0 Then
                            percentChange = (yearlyChange / startPrice) * 100
                        Else
                            percentChange = 0
                        End If
        
                        ' Update metrics for greatest increase/decrease and volume
                        If percentChange > greatestIncrease Then
                            greatestIncrease = percentChange
                            tickerGreatestIncrease = ticker
                        ElseIf percentChange < greatestDecrease Then
                            greatestDecrease = percentChange
                            tickerGreatestDecrease = ticker
                        End If
                        
                        If totalVolume > greatestVolume Then
                            greatestVolume = totalVolume
                            tickerGreatestVolume = ticker
                        End If
        
                        ' Output the data
                        With ws
                            .Cells(stockInfoRow, "I") = ticker
                            .Cells(stockInfoRow, "J") = yearlyChange
                            .Cells(stockInfoRow, "K") = percentChange & "%"
                            .Cells(stockInfoRow, "L") = totalVolume
                        End With
        
                        stockInfoRow = stockInfoRow + 1
                        totalVolume = 0
        
                        If i + 1 <= lastRow Then
                            startPrice = ws.Cells(i + 1, 3).Value
                        End If
                    End If
                Next i
        
                ' Output greatest metrics
                With ws
                    .Cells(2, "O") = "Greatest % Increase"
                    .Cells(2, "P") = tickerGreatestIncrease
                    .Cells(2, "Q") = greatestIncrease & "%"
        
                    .Cells(3, "O") = "Greatest % Decrease"
                    .Cells(3, "P") = tickerGreatestDecrease
                    .Cells(3, "Q") = greatestDecrease & "%"
        
                    .Cells(4, "O") = "Greatest Total Volume"
                    .Cells(4, "P") = tickerGreatestVolume
                    .Cells(4, "Q") = greatestVolume
        End With
    Next ws
End Sub

