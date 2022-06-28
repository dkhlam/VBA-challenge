Attribute VB_Name = "Module1"
'   this code will track the total stock volumes of the tickers in the alphabetical testing sheet

Sub ticker():
       
        ' create column names for summary table
        Cells(1, 9).Value = "Ticker"
        Cells(1, 10).Value = "Yearly Change"
        Cells(1, 11).Value = "Percent Change"
        Cells(1, 12).Value = "Total Stock Volume"

        '   variable to hold the ticker name
        tickerName = " "
        
        '   variable to hold the total stock volume of each ticker
        totalStockVolume = 0
        
        '   variable to hold the summary table starter row
        summaryTableRow = 2
        
        '   use function to find the last row in the sheet
        lastRow = Cells(Rows.Count, 1).End(xlUp).Row
        
        '   loop from row 2 in column A to the last row
        For Row = 2 To lastRow
        
            '   check to see if the ticker changes
            
                If Cells(Row + 1, 1).Value <> Cells(Row, 1).Value Then
                    '   if the ticker changes, do...
                    
                    '   first set the ticker name
                    tickerName = Cells(Row, 1).Value
                    
                    '   add the last stock volume from the row
                    totalStockVolume = totalStockVolume + Cells(Row, 7).Value
                    
                    '   add the ticker name to the I column in the summary table row
                    Cells(summaryTableRow, 9).Value = tickerName
                    
                    '   add the total stock volumes to the L column in the summary table row
                    Cells(summaryTableRow, 12).Value = totalStockVolume
                    
                    '   go to the next summary table row (add 1 on to the value of the summary table row)
                    summaryTableRow = summaryTableRow + 1
                    
                    '   reset the ticker total to 0
                    totalStockVolume = 0
                    
                Else
                    '   if the ticker stays the same, add on to the volume totals from the G column
                    totalStockVolume = totalStockVolume + Cells(Row, 7).Value
                
                End If
        
        Next Row

        
End Sub

