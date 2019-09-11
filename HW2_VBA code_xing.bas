Attribute VB_Name = "Module1"
Sub analyzeStockDataTesting()
    
For Each ws In Worksheets
    
    Dim Ratio_change As Double
    Dim open_price As Double
    Dim Close_price As Double
    
    
    
    Volume_Total = 0
    Summary_Table_Row = 2
    

    
    
    
    
    ws.Cells(1, 9) = "Ticker"
    ws.Cells(1, 10) = "Yearly Change"
    ws.Cells(1, 11) = "Percent Change"
    ws.Cells(1, 12) = "Total Stock Volume"
    
    ws.Cells(1, 16) = "Ticker"
    ws.Cells(1, 17) = "Value"
    
    ws.Cells(2, 15) = "Greatest % Increase"
    ws.Cells(3, 15) = "Greatest % Decrease"
    ws.Cells(4, 15) = "Greatest Total Volume"
    

    the_last_row = ws.Cells(Rows.Count, "A").End(xlUp).Row
    
    open_price = ws.Cells(2, 3).Value               'retrive open price for first ticker on Day 1
    Ticker_Name = ws.Cells(2, 1).Value
    
    
    For i = 2 To the_last_row
    
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
        Close_price = ws.Cells(i, 6).Value           'retrive close price for each row
        
        yearly_change = Close_price - open_price
        
        If open_price <> 0 Then
        
        
            Ratio_change = yearly_change / open_price
        
            Percent_change = Format(Ratio_change, "Percent")
        
        Else
        
            Percent_change = 0
            
        End If
        
        

        
        
        ws.Range("J" & Summary_Table_Row).Value = yearly_change
        
        ws.Range("k" & Summary_Table_Row).Value = Percent_change
        
        
        
        Ticker_Name = ws.Cells(i, 1).Value
        
        
        open_price = ws.Cells(i + 1, 3).Value           'set open price
        
        volume_of_stock = ws.Cells(i, 7).Value         'retrive volume of stock for each row when
        
        Volume_Total = Volume_Total + volume_of_stock
        
        ws.Range("I" & Summary_Table_Row).Value = Ticker_Name
        
        ws.Range("L" & Summary_Table_Row).Value = Volume_Total
        
        Summary_Table_Row = Summary_Table_Row + 1
        
        Volume_Total = 0
        
        Else
            volume_of_stock = ws.Cells(i, 7).Value        'retrive volume of stock
        
            Volume_Total = Volume_Total + volume_of_stock
        
        End If
        
    Next i
    
    
    

    

    greatest_percent_increase_ticker = ""
    greatest_percent_increase_value = 0
    
    greatest_pecent_decrease_ticker = ""
    greatest_percent_decrease_value = 0
    
    greatest_total_volume_ticker = ""
    greatest_total_volume_value = 0
    
    
    LastRowForDistinct = ws.Cells(Rows.Count, "I").End(xlUp).Row 'use ws in for loop
'    MsgBox (lastRowForDistinct)

    For n = 2 To LastRowForDistinct
    
        yearly_percent_change_value = ws.Cells(n, 11).Value
        
        yearly_change_value = ws.Cells(n, 10).Value
        
        If yearly_change_value > 0 Then
            
            ws.Cells(n, 10).Interior.ColorIndex = 4
            
        ElseIf yearly_change_value < 0 Then
            
           ws.Cells(n, 10).Interior.ColorIndex = 3
            
        End If
        
    
        If yearly_percent_change_value > greatest_percent_increase_value Then
            
            greatest_percent_increase_value = yearly_percent_change_value
            
            greatest_percent_increase_ticker = ws.Cells(n, 9).Value
            
        End If
        
        
        yearly_total_volume = ws.Cells(n, 12).Value
        
        If yearly_total_volume > greatest_total_volume_value Then
        
            greatest_total_volume_value = yearly_total_volume
            
            greatest_total_volume_ticker = ws.Cells(n, 9).Value
        
        End If
        
        
        If yearly_percent_change_value < greatest_percent_decrease_value Then
            
            greatest_percent_decrease_value = yearly_percent_change_value
            
            greatest_percent_decrease_ticker = ws.Cells(n, 9).Value
            
        End If
        
        
    Next n

    
    
    ws.Cells(2, 16).Value = greatest_percent_increase_ticker
    ws.Cells(2, 17).Value = Format(greatest_percent_increase_value, "Percent")
    
    ws.Cells(3, 16).Value = greatest_percent_decrease_ticker
    ws.Cells(3, 17).Value = Format(greatest_percent_decrease_value, "Percent")
    
    ws.Cells(4, 16).Value = greatest_total_volume_ticker
    ws.Cells(4, 17).Value = greatest_total_volume_value
    
    ' do a AutoFit  :    ws.Columns("I:L").Autofit
    
    ws.Columns("A:Q").AutoFit
    
Next ws

End Sub
