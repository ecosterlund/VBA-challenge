Sub StockMarket()

'Set a variable to cycle through the worksheets
Dim ws As Worksheet
'Start loop
    For Each ws In Worksheets
        'Create column labels for the summary table
            ws.Range("I1").Value = "Ticker"
            ws.Range("J1").Value = "Yearly Change"
            ws.Range("K1").Value = "Percent Change"
            ws.Range("L1").Value = "Total Stock Volume"
    
    'Define Variables'
    
        'Ticker Variable
        Dim ticker As String
        
        'Year Open'
        Dim year_open As Double
        year_open = ws.Cells(2, 3).Value
        
        'Year Low'
        Dim year_low As Double
        year_low = 0
        
        'Year Close'
        Dim year_close As Double
        year_close = 0
        
        'Volume'
        Dim total_stock_vol As Long
        Dim current_row_vol As Long
        total_stock_vol = 0
        current_row_vol = 0
        
        'Yearly Change'
        Dim year_change As Double
        
        'Percent Change'
        Dim percent_change As Double
        
        'Looping through the data Variables'
        Dim i As Long
        Dim lastrow As Long
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        Dim rowcount As Long 'used to display values in the chart'
        rowcount = 2 'sets starting point for rowcount'
        
        'Start of Loop'
        For i = 2 To lastrow 'used for looping through
            
            'Conditional to determine if the ticker symbol is changing
            If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
            
                ' Grab close price for current ticker
                year_close = ws.Cells(i, 6).Value
              
                'Calculate the price change for the year and move it to the summary table.
                    year_change = year_close - year_open
               
                    'Conditional if year open or year close are 0'
                    If year_open = 0 Or year_close = 0 Then
                         ws.Cells(rowcount, 11).Value = percent_change
                    Else
                         percent_change = Round((year_change / year_open * 100), 2)
                    End If
               
                'Display Percent Change'
                    ws.Cells(rowcount, 11).Value = percent_change
            
                'Display Year Change'
                    ws.Cells(rowcount, 10).Value = year_change
                    
                'Conditional to format to highlight positive or negative change.
                    If year_change >= 0 Then
                        ws.Cells(rowcount, 10).Interior.ColorIndex = 4
                    Else
                        ws.Cells(rowcount, 10).Interior.ColorIndex = 3
                    End If
                    
                'Move ticker symbol to summary table
                  ws.Cells(rowcount, 9).Value = ws.Cells(i, 1).Value
                  
                'Calculates running total volume'
                    current_row_vol = ws.Cells(i, 7).Value
                    total_stock_vol = total_stock_vol + current_row_vol
                    
                'Display Total Volume'
                    ws.Cells(rowcount, 12).Value = total_stock_vol
                    
                'Resets all values'
                    total_stock_vol = 0
                    year_open = ws.Cells(i + 1, 3).Value
                    year_close = 0
                    year_change = 0
                    percent_change = 0
                    
                'Add to Rowcount'
                    rowcount = rowcount + 1
                    
              Else
              
               'Calculates running total volume'
                current_row_vol = ws.Cells(i, 7).Value
                'total_stock_vol = current_row_vol + total_stock_vol'
                
            End If
        Next i
    Next ws
End Sub
