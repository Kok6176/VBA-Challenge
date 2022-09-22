Sub stocks_year()
'declare variable for finding last row
Dim lRow As Long
'declare variable for opening price and closing price
Dim open_price As Double
Dim close_price As Double
'declare variable for storing yearly difference
Dim yearly_diff As Double
'declare variable for calculating percentage change
Dim percentage As Double
'declare variable for calculating volume of the specific ticker
Dim vol As LongLong
'declaring variables for Bonus Work
Dim gr_increase As Double
Dim gr_decrease As Double
Dim gr_tot_vol As LongLong

'loop for each worksheet
For Each ws In Worksheets

    'finding the last row in each worksheet
    lRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    vol = 0
    j = 2
    'initializing opening price
    open_price = ws.Cells(2, 3).Value
    
    'creating the header for summary rows
    ws.Cells(1, 11).Value = "Ticker"
    ws.Cells(1, 12).Value = "Yearly Change"
    ws.Cells(1, 13).Value = "Percent Change"
    ws.Cells(1, 14).Value = "Total Stock Volume"
    
    'Header for Bonus Work
    ws.Cells(1, 18).Value = "Ticker"
    ws.Cells(1, 19).Value = "Value"
    ws.Cells(2, 17).Value = "Greatest % Increase"
    ws.Cells(3, 17).Value = "Greatest % Decrease"
    ws.Cells(4, 17).Value = "Greatest Total Volume"
    
    'initializing the variables for bonus work
    gr_increase = 0
    gr_decrease = 1
    gr_tot_vol = 0
    
    ' loop through each row
    For i = 2 To lRow
        'accumulating total stock volume
        vol = vol + ws.Cells(i, 7).Value
        'check for a break in ticker symbol
        If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
            'writing the ticker in summary table
            ws.Cells(j, 11).Value = ws.Cells(i, 1).Value
            'set year end closing price for a ticker
            close_price = ws.Cells(i, 6).Value
            'calculate yearly difference for a ticker
            yearly_diff = close_price - open_price
            'writing the yearly difference in summary table
            ws.Cells(j, 12).Value = yearly_diff
            'coloring the cell according to value
            If yearly_diff > 0 Then
                ws.Cells(j, 12).Interior.Color = RGB(0, 255, 0)
            Else
                ws.Cells(j, 12).Interior.Color = RGB(255, 0, 0)
            End If
            percentage = yearly_diff / open_price
            'writing the percentage in summary table
            ws.Cells(j, 13).Value = Format(percentage, "percent")
            'writing the Total Stock Volume in summary table
            ws.Cells(j, 14).Value = vol
            'Reset year beginning opening price for next ticker symbol
            open_price = ws.Cells(i + 1, 3).Value
            ' increment counter for summary table
            j = j + 1
            
            'Extra code for Bonus work
            
            'calculating Greatest % Increase
            If percentage > gr_increase Then
                gr_increase = percentage
                ws.Cells(2, 19).Value = Format(percentage, "percent")
                ws.Cells(2, 18).Value = ws.Cells(i, 1).Value
            End If
            'calculating Greatest % Decrease
            If percentage < gr_decrease Then
                gr_decrease = percentage
                ws.Cells(3, 19).Value = Format(percentage, "percent")
                ws.Cells(3, 18).Value = ws.Cells(i, 1).Value
            End If
            'calculating Greatest Total Volume
            If vol > gr_tot_vol Then
                gr_tot_vol = vol
                ws.Cells(4, 19).Value = vol
                ws.Cells(4, 18).Value = ws.Cells(i, 1).Value
            End If
            'reset total stock volume for next ticker symbol
            vol = 0
        End If
    Next i
    'Auto size columns
    ws.Columns("K:S").AutoFit
Next ws

End Sub