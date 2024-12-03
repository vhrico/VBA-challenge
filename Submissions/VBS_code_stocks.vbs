Sub stocks()


    ' Code logic
    
    ' stocks are pre-sorted
    ' LOOP by rows creating two seperate loops for each leaderboard
    ' IF next stock is different/ not equal (<>) to previous stock, that means we have finished our group
    ' and calculate totals
    ' ELSE, then keep summing the volume until end of row (

    ' variables
    
    Dim ws As Worksheet
    Dim stock As String
    Dim next_stock As String
    Dim volume As Double
    Dim volume_total As Double
    Dim i As Long
    Dim leaderboard_row As Long
    Dim openPrice As Double
    Dim closingPrice As Double
    Dim change As Double
    Dim pctChange As Double
    
    For Each ws In ThisWorkbook.Worksheets
        
        ' Set column headers for more profesional/standard look
        
        ws.Range("A1").Value = "Ticker"
        ws.Range("B1").Value = "Date"
        ws.Range("C1").Value = "Open"
        ws.Range("D1").Value = "High"
        ws.Range("E1").Value = "Low"
        ws.Range("F1").Value = "Close"
        ws.Range("G1").Value = "Vol"
        
        ws.Range("i1").Value = "Ticker"
        ws.Range("j1").Value = "Quaterly Change"
        ws.Range("k1").Value = "Percent Change"
        ws.Range("l1").Value = "Total Stock Volume"
        ws.Range("O1").Value = "Ticker"
        ws.Range("P1").Value = "Value"
        ws.Range("N2").Value = "Greatest % Increase"
        ws.Range("N3").Value = "Greatest % Decrease"
        ws.Range("N4").Value = "Greatest Total Volume"
    

    ' Reset per stock
    volume_total = 0
    open_price = ws.Cells(2, 3).Value
    leaderboard_row = 2
   

    For i = 2 To 93001 ' Begining of First loop
        
        ' extract values from workbook
        stock = ws.Cells(i, 1).Value
        volume = ws.Cells(i, 7).Value
        next_stock = ws.Cells(i + 1, 1).Value

        ' if statement
        If (stock <> next_stock) Then
            
            ' add total
            volume_total = volume_total + volume
            closing_price = ws.Cells(i, 6).Value
            change = closing_price - open_price
            pct_change = change / open_price
            
            ' write to leaderboard
            ws.Cells(leaderboard_row, 12).Value = volume_total
            ws.Cells(leaderboard_row, 11).Value = FormatPercent(pct_change)
            ws.Cells(leaderboard_row, 10).Value = change
            ws.Cells(leaderboard_row, 9).Value = stock
            
            ' Conditional Formatting will change the color of
            ' 'Quarterly change' (change variable) if positive (green= 4) or negative (red=3)
            If (change > 0) Then
                ws.Cells(leaderboard_row, 10).Interior.ColorIndex = 4
            ElseIf (change < 0) Then
                ws.Cells(leaderboard_row, 10).Interior.ColorIndex = 3
            Else
                ' Do Nothing
            End If



            ' reset total
            volume_total = 0
            leaderboard_row = leaderboard_row + 1
            open_price = ws.Cells(i + 1, 3).Value ' open price of NEXT stock
        Else
            ' add total
            volume_total = volume_total + volume
        End If
    Next i ' Returns to start of loop
    
    ' Second Loop for Second Leaderboard
    Dim max_price As Double
    Dim min_price As Double
    Dim max_volume As LongLong
    Dim max_price_stock As String
    Dim min_price_stock As String
    Dim max_volume_stock As String
    Dim j As Integer
    
    
    max_price = ws.Cells(2, 11).Value
    min_price = ws.Cells(2, 11).Value
    max_volume = ws.Cells(2, 12).Value
    max_price_stock = ws.Cells(2, 9).Value
    min_price_stock = ws.Cells(2, 9).Value
    max_volume_stock = ws.Cells(2, 9).Value
    
    
    ' This loop will 'look' for new max and min percent change on value of every stock
    ' every new max or min change will be assigned to its corresponding variable (max_price, min_price)
    ' Until all thats left is the max and min value for that 'Percent change' column
    
    For j = 2 To leaderboard_row
        If (ws.Cells(j, 11).Value > max_price) Then
            ' new max price change
            max_price = ws.Cells(j, 11).Value
            max_price_stock = ws.Cells(j, 9).Value
        End If
        
        If (ws.Cells(j, 11).Value > max_price) Then
            ' new min price change
            min_price = ws.Cells(j, 11).Value
            min_price_stock = ws.Cells(j, 9).Value
        End If
        
        If (ws.Cells(j, 12).Value > max_volume) Then
            ' new max volume change
            max_volume = ws.Cells(j, 12).Value
            max_volume_stock = ws.Cells(j, 9).Value
        End If
        
    Next j
    
    ' Write out values to Excel Workbook
    ws.Range("O2").Value = max_price_stock
    ws.Range("O3").Value = min_price_stock
    ws.Range("O4").Value = max_volume_stock
    
    ws.Range("P2").Value = FormatPercent(max_price) ' Formats value to percent
    ws.Range("P3").Value = FormatPercent(min_price) ' Formats value to percent
    ws.Range("P4").Value = max_volume
    
    Next ws

End Sub

