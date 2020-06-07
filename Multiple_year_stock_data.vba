Sub stock_data()
    
Dim sht As Worksheet
For Each sht In Worksheets
    
    'Create Headers for each worksheet
    sht.Cells(1, 9).Value = "Ticker"
    sht.Cells(1, 10).Value = "Yearly Change"
    sht.Cells(1, 11).Value = "Percent Change"
    sht.Cells(1, 12).Value = "Total Stock Volume"
    
    j = 2
    LastRow = sht.Cells(Rows.Count, 1).End(xlUp).Row
    
    Dim StockOpen As Double
    Dim StockClose As Double
    Dim YearlyChange As Double
    StockOpen = 0
        
    For i = 2 To LastRow

        'Let us calculate Total Stock Volume by adding next value in the column to the previous one
        Dim TotalStockVolume As Double
        TotalStockVolume = TotalStockVolume + sht.Cells(i, 7)
        
        'Now we have to capture the value of StockOpen for each new stock.
        'One way of doing this is to check the new ticker symbol to the current one.
        'And as soon as it doesn't match, we would store the value for <open> at that cell in StockOpen
        If sht.Cells(i, 1).Value <> sht.Cells(i - 1, 1).Value Then
            StockOpen = sht.Cells(i, 3).Value
        End If
    
        'Let us populate the cells that we just created
        'Now here, we use the same concept as above
        If sht.Cells(i, 1).Value <> sht.Cells(i + 1, 1).Value Then
            'This will fill the ticker symbol
            sht.Cells(j, 9).Value = sht.Cells(i, 1).Value
            'This will fill the Total Stock Volume that we calculated above
            sht.Cells(j, 12).Value = TotalStockVolume
            'Closing stock value will be the values in <close> column
            StockClose = sht.Cells(i, 6).Value
            'Yearly change will be the (closing stock - opening stock(which we calculated above))
            'Let us store this in a variable so that we can reuse it and set it to 0 for the next stock
            YearlyChange = StockClose - StockOpen
            sht.Cells(j, 10).Value = YearlyChange
        
            'Highlight positive change in green and negative change in red
            If YearlyChange >= 0 Then
                sht.Cells(j, 10).Interior.Color = vbGreen
            Else
                sht.Cells(j, 10).Interior.Color = vbRed
            End If
        
            'Let us calculate percent change but if the opening and closing stock are both 0, then it results in error.
            'So let us first create an If statement to catch this.
            'Also, this needs to be displayed in % format.
            'One way to do it is by using the .NumberFormat function and setting it to show 2 decimal places.
            If StockOpen = 0 Or StockClose = 0 Then
                PercentChange = 0
                sht.Cells(j, 11).Value = PercentChange
                sht.Cells(j, 11).NumberFormat = "0.00%"
            Else
                PercentChange = YearlyChange / StockOpen
                sht.Cells(j, 11).Value = PercentChange
                sht.Cells(j, 11).NumberFormat = "0.00%"
            End If
        
            'Increment j by 1
            j = j + 1
        
            'Reset all values to 0
            StockOpen = 0
            StockClose = 0
            TotalStockVolume = 0
            PercentChange = 0
            YearlyChange = 0
        End If
        Next i
        
        'CHALLENGES
        'Let us make the challenge table
        sht.Cells(2, 15).Value = "Greatest % Increase"
        sht.Cells(3, 15).Value = "Greatest % Decrease"
        sht.Cells(4, 15).Value = "Greatest Total Volume"
        sht.Cells(1, 16).Value = "Ticker"
        sht.Cells(1, 17).Value = "Value"

        LastRowTicker = sht.Cells(Rows.Count, 9).End(xlUp).Row
        
        'Let us set the value of first stock to the first value in Percent Change column
        'So that will become our base value and then we can compare subsequent values to that in order to find the Greatest Increase
        ValueHigh = sht.Cells(2, 11).Value
        'Let us do the same what we did above to find the Greatest Decrease
        ValueLow = sht.Cells(2, 11).Value
        StockVolume = sht.Cells(2, 12).Value
        
        For j = 2 To LastRowTicker
            'Now let us find the Greatest Increase
            If sht.Cells(j, 11).Value > ValueHigh Then
                ValueHigh = sht.Cells(j, 11).Value
                TickerHigh = sht.Cells(j, 9).Value
            End If
            
            'And let us find the Greatest Decrease
            If sht.Cells(j, 11).Value < ValueLow Then
                ValueLow = sht.Cells(j, 11).Value
                TickerLow = sht.Cells(j, 9).Value
            End If
            
            'Here we will find the Greatest Total Volume of the highest ticker
            If sht.Cells(j, 12).Value > StockVolume Then
                StockVolume = sht.Cells(j, 12).Value
                TickerTop = sht.Cells(j, 9).Value
            End If
        Next j
        
        'Now let us populate the challenge table with all this information
        sht.Cells(2, 16).Value = TickerHigh
        sht.Cells(2, 17).Value = ValueHigh
        sht.Cells(2, 17).NumberFormat = "0.00%"
        sht.Cells(3, 16).Value = TickerLow
        sht.Cells(3, 17).Value = ValueLow
        sht.Cells(3, 17).NumberFormat = "0.00%"
        sht.Cells(4, 16).Value = TickerTop
        sht.Cells(4, 17).Value = StockVolume
    Next sht
End Sub