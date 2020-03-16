' Stock Wolf v1.0 VBA script for analyzing individual stocks for the year
' 1. Consolidates individual ticker symbols into 1 list
' 2. Finds change from opening to closing price of year
' 3. Gets total volume for the year per ticker symbol
' 4. Calculates greatest increase and decrease, and volume for year 

Sub StockWolf()

    ' --------------------------------------------
    ' LOOP THROUGH ALL SHEETS PER YEAR
    ' --------------------------------------------
    For Each ws In Worksheets

        ' Create variables for yearly calculations per symbol
        Dim Ticker, row As Integer
        Dim openPrice, closePrice, yearChange As Double
        Dim volume As Double
        ' Init variables for Ticker symbol column and row to start for output, init volume to zero
        Ticker = 1
        row = 2
        volume = 0

        ' Determine the Last Row
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).row
        
        ' Get very first symbol in the sheet
        currentSymbol = ws.Cells(2, Ticker).Value
        ' Grab opening price for very first symbol
        openPrice = ws.Cells(2, 3).Value
        
        ' ---------------------------------------------
        ' SET UP WORKSHEET LABELS, ETC FOR FINAL OUTPUT
        ' ---------------------------------------------
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percentage Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
            
        ' Loop through rows in the column of stock symbols
        For i = 2 To LastRow
            ' Calculate and print output for current symbol if next one is a new symbol
            If ws.Cells(i + 1, Ticker).Value <> ws.Cells(i, Ticker).Value Then
                ' Print current symbol, get close price for last day of year, and calculate year change and print
                ws.Cells(row, 9).Value = currentSymbol
                closePrice = ws.Cells(i, 6).Value
                yearChange = closePrice - openPrice
                ws.Cells(row, 10).Value = yearChange
                    ' Calculate percentage change and account for divide by zero
                    If openPrice <= 0 Then
                        ws.Cells(row, 11).Value = 0
                    Else
                        ws.Cells(row, 11).Value = yearChange / openPrice
                    End If
                ' Format output for percentages
                ws.Cells(row, 10).NumberFormat = "0.00"
                ws.Cells(row, 11).NumberFormat = "0.00%"
                ' Add last row for symbols volume to total and print
                volume = volume + ws.Cells(i, 7)
                ws.Cells(row, 12).Value = volume
                ' Reinitialize for next symbol with open price and reset volume and row for output
                currentSymbol = ws.Cells(i + 1, Ticker).Value
                openPrice = ws.Cells(i + 1, 3).Value
                row = row + 1
                volume = 0
            End If
                ' If next symbol is the same, add the current rows volume to running total  
                volume = volume + ws.Cells(i, 7)

        Next i
        
        ' Determine how many rows in column of consolidated symbols
        LastTicker = ws.Cells(Rows.Count, 9).End(xlUp).row
        
        ' Loop through percentage chanes and color green or red according to increase or decrease
        For i = 2 To LastTicker
            If ws.Cells(i, 10).Value >= 0 Then
                ws.Cells(i, 10).Interior.ColorIndex = 4
            Else
                ws.Cells(i, 10).Interior.ColorIndex = 3
            End If
        Next i
        
        ' Create variables to hold min, max, and max volume calculations and ticker symbols
        Dim maxPercent, minPercent, maxVolume As Double
        Dim maxPercentTicker, minPercentTicker, maxVolumeTicker As String
        
        ' Initialize variables to very first symbol in column
        maxPercent = ws.Cells(2, 11).Value
        minPercent = ws.Cells(2, 11).Value
        maxVolume = ws.Cells(2, 12).Value
        maxPercentTicker = ws.Cells(2, 9).Value
        minPercentTicker = ws.Cells(2, 9).Value
        maxVolumeTicker = ws.Cells(2, 9).Value
        
        ' Loop through output and find min, max, max volume for entire list
        For i = 2 To LastTicker
            If ws.Cells(i + 1, 11).Value > maxPercent Then
                maxPercent = ws.Cells(i + 1, 11).Value
                maxPercentTicker = ws.Cells(i + 1, 9).Value
            End If
            If ws.Cells(i + 1, 11).Value < minPercent Then
                minPercent = ws.Cells(i + 1, 11).Value
                minPercentTicker = ws.Cells(i + 1, 9).Value
            End If
            If ws.Cells(i + 1, 12).Value > maxVolume Then
                maxVolume = ws.Cells(i + 1, 12).Value
                maxVolumeTicker = ws.Cells(i + 1, 9).Value
            End If
        Next i
        
        ' Fill cells with final calculations per sheet and format, autofit
        ws.Cells(2, 16).Value = maxPercentTicker
        ws.Cells(3, 16).Value = minPercentTicker
        ws.Cells(4, 16).Value = maxVolumeTicker
        ws.Cells(2, 17).Value = maxPercent
        ws.Cells(3, 17).Value = minPercent
        ws.Cells(4, 17).Value = maxVolume

        ws.Cells(2, 17).NumberFormat = "0.00%"
        ws.Cells(3, 17).NumberFormat = "0.00%"

        ws.Columns("A:Q").AutoFit

    ' Go to next sheet 
    Next ws

End Sub
