Sub Stockloop()
        
    Dim yearChange As Double
    Dim pctChange As Double
    Dim nextTick As String
    Dim currTick As String
    Dim curropenprice As Double
    Dim nextopenprice As Double
    Dim totalVol As Variant
    Dim ws As Worksheet
    Dim rowCount As Long
    Dim pricechangecells As Variant
    Dim volumecells As Variant
    Dim maxchange As Double
    Dim minchange As Double
    Dim maxvolume As Double
    
    Dim summTableRow As Integer
    summTableRow = 2
    
    
    'Loop through all worksheets
    For Each ws In Worksheets
        
        ws.Activate
        
        'Get total row count for the worksheet
        rowCount = Worksheets("ws.name").Cells.SpecialCells(xlCellTypeLastCell).Row
        
        'Init Summary Table Headers
        Cells(1, 9).Value = "Ticker"
        Cells(1, 10).Value = "Yearly Change"
        Cells(1, 11).Value = "Percent Change"
        Cells(1, 12).Value = "Total Stock Volume"
        
        'Init current ticker
        currTick = Cells(2, 1).Value
        
        'Init open price
        curropenprice = Cells(2, 3).Value
        
        
            'Loop thorugh all rows
            For i = 1 To rowCount
                
                'Check Look Ahead next tick
                nextTick = Cells(i + 1, 1).Value
                
                If nextTick <> currTick Then
                    
                    'Calculate stats and write to our summary table
                    yearChange = Cells(i, 6) - curropenprice
                    pctChange = (Cells(i, 6) - curropenprice) / curropenprice
                    
                    'Post Ticker to summary table
                    Cells(summTableRow, 9).Value = currTick
                    
                    'Post Yearly Change to table
                    Cells(summTableRow, 10).Value = yearChange
                    
                    'Post pct Change to table
                    Cells(summTableRow, 11).Value = pctChange
                    
                    'update open price
                    nextopenprice = Cells(i + 1, 3).Value
                    
                    'Add days volume to total volume
                    totalVol = totalVol + Cells(i, 7).Value
                    
                    'Post total volume to summary table
                    Cells(summTableRow, 12).Value = totalVol
                    
                    'Reset Total volume
                    totalVol = 0
                    
                    'move to next symbol
                    currTick = nextTick
                    summTableRow = summTableRow + 1
                    curropenprice = nextopenprice
                    
                    
                'Else add days volume to total volume
                Else
                    totalVol = totalVol + Cells(i, 7).Value
                    
                End If
            
            Next i
            
            'Set cell formatting for price change cell color
            
            'get count of number of price change cells to color
            pricechangecells = ws.Cells(Rows.Count, "J").End(xlUp).Row
            
            For i = 1 To pricechangecells
            
                'Green for positive price change
                If Cells(i, 10).Value > 0 Then
                    Cells(i, 10).Interior.ColorIndex = 4
                End If
                
                'Red for negative price change
                If Cells(i, 10).Value < 0 Then
                    Cells(i, 10).Interior.ColorIndex = 3
                End If
        
            Next i
            
            
            'Get largest pct return, smallest pct return and largest volume
            
            'init row and column titles
            Cells(2, 15).Value = "Greatest % Increase"
            Cells(3, 15).Value = "Greatest % Decrease"
            Cells(4, 15).Value = "Greatest Total Volume"
            Cells(1, 16).Value = "Ticker"
            Cells(1, 17).Value = "Value"
            
            'set percent change and volume ranges
            'Didn't know how to specify a range dependend upon cell having a value so used 10,000
            pctchangerange = Range("K2:K10000")
            volumerange = Range("L2:L10000")
            
            'Determine max change, max loss, max vol in their respective ranges
            maxchange = Application.WorksheetFunction.Max(pctchangerange)
            minchange = Application.WorksheetFunction.Min(pctchangerange)
            maxvolume = Application.WorksheetFunction.Max(volumerange)
            
            
            'Loop through price change cells to gather max change and min change and tickers
            
            'Determine range of price change cells by finding end row
            pctchangecells = ws.Cells(Rows.Count, "K").End(xlUp).Row
            
            For i = 2 To pctchangecells
                If Cells(i, 11).Value = maxchange Then
                
                    'post value of stock ticker to table
                    Cells(2, 16).Value = Cells(i, 10).Value
                    
                    'post value of maxchange cell to table
                    Cells(2, 17).Value = Cells(i, 11).Value
                    
                ElseIf Cells(i, K).Value = minchange Then
                    
                    'post value of stock ticker to table
                    Cells(3, 16).Value = Cells(i, 10).Value
                    
                    'post value of minchange cell to table
                    Cells(3, 17).Value = Cells(i, 11).Value
                End If
                
            Next i
            
                        
            'Loop through volume cells to gather highest volume stock
            
            'Determine range of volume cells by finding end row
            volumecells = ws.Cells(Rows.Count, "L").End(xlUp).Row
            
            For i = 2 To volumecells
                If Cells(i, 12).Value = maxvolume Then
                
                    'post value of largest volume stock ticker to table
                    Cells(4, 16).Value = Cells(i, 12).Value
                    
                    'post value of largest volume stock to table
                    Cells(4, 17).Value = Cells(i, 12).Value
                End If
                
            Next i
                                   
    Next ws
    
End Sub

