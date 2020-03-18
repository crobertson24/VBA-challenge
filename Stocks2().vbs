Sub Stocks2():

Dim Ticker As String

Dim lastRowState As Long
Dim openingPrice As Double
Dim closingPrice As Double
Dim yearlyChange As Double
Dim percentChange As Double
Dim totalStockVolume As Double
Dim greatestPercentIncrease As Double
Dim greatestPercentDecrease As Double
Dim numberTickers As Integer
Dim greatestStockVolume As Double
Dim greatestPercentIncreaseTicker As String
Dim greatestPercentDecreaseTicker As String
Dim greatestStockVolumeTicker As String

For Each ws In Worksheets

    ws.Activate

    lastRowState = ws.Cells(Rows.Count, "A").End(xlUp).Row

    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    
    numberTickers = 0
    Ticker = ""
    yearlyChange = 0
    openingPrice = 0
    percentChange = 0
    totalStockVolume = 0
    
    For i = 2 To lastRowState

        Ticker = Cells(i, 1).Value
        
        If openingPrice = 0 Then
            openingPrice = Cells(i, 3).Value
        End If
        
        totalStockVolume = totalStockVolume + Cells(i, 7).Value
        
        If Cells(i + 1, 1).Value <> Ticker Then
            numberTickers = numberTickers + 1
            Cells(number_tickers + 1, 9) = Ticker
            
            closingPrice = Cells(i, 6)
            
            yearlyChange = closingPrice - openingPrice
            
            Cells(numberTickers + 1, 10).Value = yearlyChange
            
            If yearlyChange > 0 Then
                Cells(numberTickers + 1, 10).Interior.ColorIndex = 4

            ElseIf yearlyChange < 0 Then
                Cells(numberTickers + 1, 10).Interior.ColorIndex = 3
        
            Else
                Cells(numberTickers + 1, 10).Interior.ColorIndex = 6
            End If
            
            
            If openingPrice = 0 Then
                percentChange = 0
            Else
                percentChange = (yearlyChange / openingPrice)
            End If
            
            Cells(numberTickers + 1, 11).Value = Format(percentChange, "Percent")
            
            
           
            If percent_change > 0 Then
                 Cells(number_tickers + 1, 11).Interior.ColorIndex = 4
     
            ElseIf percent_change < 0 Then
                Cells(number_tickers + 1, 11).Interior.ColorIndex = 3
           
            Else
                Cells(number_tickers + 1, 11).Interior.ColorIndex = 6
            End If
            
            
            openingPrice = 0
            
            Cells(numberTickers + 1, 12).Value = totalStockVolume
            
            totalStockVolume = 0
        End If
        
    Next i
    
    Range("O2").Value = "Greatest % Increase"
    Range("O3").Value = "Greatest % Decrease"
    Range("O4").Value = "Greatest Total Volume"
    Range("P1").Value = "Ticker"
    Range("Q1").Value = "Value"
    
    
    lastRowState = ws.Cells(Rows.Count, "I").End(xlUp).Row
    

    greatestPercentIncrease = Cells(2, 11).Value
    greatestPercentDecrease = Cells(2, 11).Value
    greatestStockVolume = Cells(2, 12).Value
    greatestPercentIncreaseTicker = Cells(2, 9).Value
    greatestPercentDecreaseTicker = Cells(2, 9).Value
    greatestStockVolumeTicker = Cells(2, 9).Value
    
    
    For i = 2 To lastRowState
    
    
        If Cells(i, 11).Value > greatestPercentIncrease Then
            greatestPercentIncrease = Cells(i, 11).Value
            greatestPercentIncreaseTicker = Cells(i, 9).Value
        End If
        
        If Cells(i, 11).Value < greatestPercentDecrease Then
            greatestPercentDecrease = Cells(i, 11).Value
            greatestPercentDecreaseTicker = Cells(i, 9).Value
        End If
        
        If Cells(i, 12).Value > greatestStockVolume Then
            greatestStockVolume = Cells(i, 12).Value
            greatestStockVolumeTicker = Cells(i, 9).Value
        End If
        
    Next i
    
    Range("Q2").Value = Format(greatestPercentIncrease, "Percent")
    Range("Q3").Value = Format(greatestPercentDecrease, "Percent")
    Range("Q4").Value = greatest_StockVolume
    Range("P2").Value = Format(greatestPercentIncreaseTicker, "Percent")
    Range("P3").Value = Format(greatestPercentDecreaseTicker, "Percent")
    Range("P4").Value = greatest_Stock_VolumeTicker
    
    
Next ws


End Sub
