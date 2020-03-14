Sub StockWolf()

    ' --------------------------------------------
    ' LOOP THROUGH ALL SHEETS
    ' --------------------------------------------
    For Each ws In Worksheets

        ' --------------------------------------------
        ' INSERT THE STATE
        ' --------------------------------------------

        ' Created a Variable to Hold File Name, Last Row, Last Column, and Year
        'Dim WorksheetName As String
        Dim Ticker, row As Integer
        Dim openPrice, closePrice As Double
        Dim volume As Double
        Ticker = 1
        row = 2
        volume = 0
        openPrice = ws.Cells(2, 3).Value

        ' Determine the Last Row
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).row
        
        currentSymbol = ws.Cells(2, Ticker).Value
        
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percentage Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
  ' Loop through rows in the column
  For i = 2 To LastRow
    
    If ws.Cells(i + 1, Ticker).Value <> ws.Cells(i, Ticker).Value Then
        ws.Cells(row, 9).Value = currentSymbol
        closePrice = ws.Cells(i, 6).Value
        ws.Cells(row, 10).Value = closePrice - openPrice
        ws.Cells(row, 11).Value = ws.Cells(row, 10) / openPrice
        ws.Cells(row, 10).NumberFormat = "0.00"
        ws.Cells(row, 11).NumberFormat = "0.00%"
        volume = volume + ws.Cells(i, 7)
        ws.Cells(row, 12).Value = volume
        currentSymbol = ws.Cells(i + 1, Ticker).Value
        openPrice = ws.Cells(i + 1, 3).Value
        row = row + 1
        volume = 0
    End If
        volume = volume + ws.Cells(i, 7)
  Next i
  
  LastTicker = ws.Cells(Rows.Count, 9).End(xlUp).row
  For i = 2 To LastTicker
     If ws.Cells(i, 10).Value >= 0 Then
        ws.Cells(i, 9).Interior.ColorIndex = 4
     Else
        ws.Cells(i, 9).Interior.ColorIndex = 3
     End If
  Next i
  
Next ws

End Sub
