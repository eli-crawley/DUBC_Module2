Sub quarterlyStocks()

    'create variables for stock ticker
    Dim ticker_symbol As String
    
    'create a variable for holding the opening price
    Dim open_price As Double

    'create a variable for holding the closing price
    Dim close_price As Double
    
    'create a variable for holding the price change
    Dim price_change As Double
    
    'create a variable for percent change
    Dim percent_change As Double
    
    'create a varialbe for holding the total stock volume
    Dim stock_volume As Double
    
    'create variables for maximum percent increase, max percent decrease and max volume
    Dim max_percent_increase As Double
    Dim max_percent_decrease As Double
    Dim max_volume As Double
    
    'create variables for holding the ticker symbols for the ma variables
    Dim max_percent_increase_ticker As String
    Dim max_percent_decrease_ticker As String
    Dim max_volume_ticker As String
    
    'Set variables for loop
    Dim i As Long
    Dim LastRow As Long

    'Keep track of the location for ticker in the summary table
    Dim summary_row As Long
    Dim ws As Worksheet
        
    'Loop through the tickers in each worksheet
    For Each ws In Worksheets
    
        'Initialize the summary row
        summary_row = 2
        
        'determine the last row
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        'Reset max values as the start of the worksheet
        max_percent_increase = 0
        max_percent_decrease = 0
        max_volume = 0
        max_percent_increase_ticker = ""
        max_percent_decrease_ticker = ""
        
        'Loop through ticker symbol
        For i = 2 To LastRow
    
            'Check if we are at the first row of the ticker
            If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then
        
                'set the ticker symbol
                ticker_symbol = ws.Cells(i, 1).Value
            
                'set opening price using the first row
                open_price = ws.Cells(i, 3).Value
            
                'initialize stock volume for the ticker
                stock_volume = 0
            
            End If
        
            'add to the stock volume total
            stock_volume = stock_volume + ws.Cells(i, 7).Value
        
            'Check if we are at the last line of the ticker
            If i = LastRow Or ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                        
                'set the closing price
                close_price = ws.Cells(i, 6).Value
            
                'set price change
                price_change = close_price - open_price
                
                'print the ticker symbol in the summary table
                ws.Range("I" & summary_row).Value = ticker_symbol
     
                'print the quarterly change in stock price in the summary table
                ws.Range("J" & summary_row).Value = price_change
                
                'set percent change and round to two decimal places
                If open_price <> 0 Then
                    percent_change = Round((price_change / open_price) * 100, 2)
                Else
                    percent_change = 0
                End If
                                  
                'print the percent change in the stock price and round to 2 decimal places
                ws.Range("K" & summary_row).Value = percent_change
                            
                'print the stock volume total in the summary table
                ws.Range("L" & summary_row).Value = stock_volume
            
                'add conditional formatting for price_change
                If price_change > 0 Then
                    ws.Range("J" & summary_row).Interior.ColorIndex = 4
                ElseIf price_change < 0 Then
                    ws.Range("J" & summary_row).Interior.ColorIndex = 3
                End If
            
                'add on the the summary row
                summary_row = summary_row + 1
                               
                'check and update max increase values
                If percent_change > max_percent_increase Then
                    max_percent_increase = percent_change
                    max_percent_increase_ticker = ticker_symbol
                End If
                
                'check and update max decrease values
                If percent_change < max_percent_decrease Then
                    max_percent_decrease = percent_change
                    max_percent_decrease_ticker = ticker_symbol
                End If
                
                'check and update max volume
                If stock_volume > max_volume Then
                    max_volume = stock_volume
                    max_volume_ticker = ticker_symbol
                End If
                
            End If
        
        Next i
        
        'print max vlues for each worksheet
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("P2").Value = max_percent_increase_ticker
        ws.Range("Q2").Value = max_percent_increase
        
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("P3").Value = max_percent_decrease_ticker
        ws.Range("Q3").Value = max_percent_decrease
        
        ws.Range("O4").Value = "Greatest Total Volume"
        ws.Range("P4").Value = max_volume_ticker
        ws.Range("Q4").Value = max_volume
 
    
    Next ws
    
End Sub

