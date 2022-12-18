Attribute VB_Name = "Module1"
Sub stocks()

'PART 1: LOOP THROUGH ALL THE STOCKS

For Each ws In Worksheets

    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    
    Dim Worksheet As String
    Dim lastrow, column, row As Integer
    Dim ticker As String
    Dim yearly_change, percent_change, total_stock_volume As Double
    Dim opening, closing As Double
    Dim lastrowsummary, rowsummary, columnsummary As Integer
    Dim ticker_found_1, ticker_found_2, ticker_found_3 As String
    Dim greatest_increase, greatest_decrease, greatest_volume As Double
            
    Worksheet = ws.Name
    
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).row
    row = 2
    column = 1

    opening = ws.Cells(row, column + 2).Value
    total_stock_volume = 0

    For i = 2 To lastrow
    
        If ws.Cells(i + 1, column).Value <> ws.Cells(i, column).Value Then
            
            'GET THE TICKER SYMBOL
            
            ticker = ws.Cells(i, column).Value
            ws.Cells(row, 9).Value = ticker
            
            'GET THE YEARLY PRICE CHANGE
            
            closing = ws.Cells(i, column + 5).Value
            yearly_change = closing - opening
            ws.Cells(row, 10).Value = yearly_change
        
                If yearly_change > 0 Then
                    ws.Cells(row, 10).Interior.ColorIndex = 4
                Else
                    ws.Cells(row, 10).Interior.ColorIndex = 3
                End If
            
            'GET THE PERCENTAGE OF PRICE CHANGE
            
            percent_change = (closing - opening) / opening
            ws.Cells(row, 11).Value = percent_change
            ws.Cells(row, 11).NumberFormat = "0.00%"
            
            ' GET THE TOTAL VOLUME OF THE STOCK
            
            total_stock_volume = total_stock_volume + ws.Cells(i, column + 6)
            ws.Cells(row, 12).Value = total_stock_volume
        
            row = row + 1
            opening = ws.Cells(i + 1, column + 2).Value
            total_stock_volume = 0
        
        Else
    
            total_stock_volume = total_stock_volume + ws.Cells(i, column + 6)
        
        End If
    
    Next i
    
'PART 2: BONUS ADD FUNCTIONALITY TO THE SCRIPT
    
    lastrowsummary = ws.Cells(Rows.Count, 9).End(xlUp).row
    rowsummary = 2
    columnsummary = 9
        
    ws.Cells(1, 15).Value = "Summary"
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"
    
    greatest_increase = ws.Cells(rowsummary, columnsummary + 2).Value
    greatest_decrease = ws.Cells(rowsummary, columnsummary + 2).Value
    greatest_volume = ws.Cells(rowsummary, columnsummary + 3).Value
    ticker_found_1 = ws.Cells(rowsummary, columnsummary).Value
    ticker_found_2 = ws.Cells(rowsummary, columnsummary).Value
    ticker_found_2 = ws.Cells(rowsummary, columnsummary).Value
                
    For j = 2 To lastrowsummary
        
        'FIND GREATEST PERCENTAGE OF INCREASE AND ITS TICKER SYMBOL
        
        If ws.Cells(j + 1, columnsummary + 2).Value > greatest_increase Then
            ticker_found_1 = ws.Cells(j + 1, columnsummary).Value
            greatest_increase = ws.Cells(j + 1, columnsummary + 2).Value
                    
        Else
            ticker_found_1 = ticker_found_1
            greatest_increase = greatest_increase
                        
        End If
        
        ws.Cells(2, 16).Value = ticker_found_1
        ws.Cells(2, 17).Value = greatest_increase
        ws.Cells(2, 17).NumberFormat = "0.00%"
        
        'FIND GREATEST PERCENTAGE OF DECREASE AND ITS TICKER SYMBOL
        
        If ws.Cells(j + 1, columnsummary + 2).Value < greatest_decrease Then
            ticker_found_2 = ws.Cells(j + 1, columnsummary).Value
            greatest_decrease = ws.Cells(j + 1, columnsummary + 2).Value
                    
        Else
            ticker_found_2 = ticker_found_2
            greatest_decrease = greatest_decrease
                        
        End If
        
        ws.Cells(3, 16).Value = ticker_found_2
        ws.Cells(3, 17).Value = greatest_decrease
        ws.Cells(3, 17).NumberFormat = "0.00%"
        
        'FIND GREATEST TOTAL VOLUME AND ITS TICKER SYMBOL
        
        If ws.Cells(j + 1, columnsummary + 3).Value > greatest_volume Then
            ticker_found_3 = ws.Cells(j + 1, columnsummary).Value
            greatest_volume = ws.Cells(j + 1, columnsummary + 3).Value
                    
        Else
            ticker_found_3 = ticker_found_3
            greatest_volume = greatest_volume
                        
        End If
                                           
        ws.Cells(4, 16).Value = ticker_found_3
        ws.Cells(4, 17).Value = greatest_volume
        
    Next j
                
    ws.Range("I:R").EntireColumn.AutoFit
       
Next ws
  
End Sub
