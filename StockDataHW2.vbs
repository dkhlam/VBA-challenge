Attribute VB_Name = "Module1"
Sub stockData():

        Dim totalStockVolume As Double  '   total stock volume
        Dim row As Long     ' loop control variable that will go through the rows of the sheet
        Dim rowCount As Long    '  variable for number of rows in the sheet
        Dim Change As Double ' variable for yearly change for each stock in the sheet
        Dim percentChange As Double  ' variable that holds the percent change for each stock in the sheet
        Dim summaryTableRow As Long ' variable that holds the rows of the summary table row
        Dim stockStartRow As Long ' variable that holds the start of a stock's rows in the sheet
       
       ' loop through all the worksheets
       For Each ws In Worksheets
       
                ' create column names for summary table
                ws.Cells(1, 9).Value = "Ticker"
                ws.Cells(1, 10).Value = "Yearly Change"
                ws.Cells(1, 11).Value = "Percent Change"
                ws.Cells(1, 12).Value = "Total Stock Volume"
                
                ' auto fit the column names
                ws.Columns("A:Q").AutoFit
                
                ' initialize the values
                summaryTableRow = 0     ' summary table row starts at 0
                totalStockVolume = 0   ' total stock volume for a stock starts at 0
                yearlyChange = 0    ' yearly change starts at 0
                stockStartRow = 2   ' first stock in the sheet starts on row
                
                '   use function to get value of last row in the sheet
                rowCount = ws.Cells(Rows.Count, 1).End(xlUp).row
        
                '   loop from row 2 in column A to the last row
                For row = 2 To rowCount
                
                    '   check to see if ticker changes
                    If ws.Cells(row + 1, 1).Value <> ws.Cells(row, 1).Value Then
                    
                        '    calculate the total one last time for the ticker
                        totalStockVolume = totalStockVolume + ws.Cells(row, 7).Value ' grabs stock volume from column G
                        
                        '   check to see if the total volume is 0
                        If totalStockVolume = 0 Then
                            '   print results in summary table
                            ws.Range("I" & 2 + summaryTableRow).Value = ws.Cells(row, 1).Value '  prints the ticker name in Column I
                            ws.Range("J" & 2 + summaryTableRow).Value = 0 '  prints the 0 in Column J ' prints the 0 in column J (yearly change)
                            ws.Range("K" & 2 + summaryTableRow).Value = 0 & "%"  '  prints the 0% in Column K (percent change)
                            ws.Range("L" & 2 + summaryTableRow).Value = 0 '  prints the 0 in Column L (total stock volume)
                        Else
                            '   find the first non zero starting value
                            If ws.Cells(stockStartRow, 3).Value = 0 Then
                                For findValue = Start To row
                                
                                    '   check to see if the next value does not equal 0
                                    If ws.Cells(findValue, 3).Value <> 0 Then
                                        stockStartRow = findValue
                                        '   once value is non-zero value, break out of loop
                                        Exit For
                                    End If
                                Next findValue
                            End If
                            
                            ' Calculate the yearly change (last close - first open)
                            yearlyChange = (ws.Cells(row, 6).Value - ws.Cells(stockStartRow, 3).Value)
                            ' Calculate the percent change (yearly change / first open)
                            percentChange = yearlyChange / ws.Cells(stockStartRow, 3).Value
                            
                            '   print results in summary table
                            ws.Range("I" & 2 + summaryTableRow).Value = ws.Cells(row, 1).Value '  prints the ticker name in Column I
                            ws.Range("J" & 2 + summaryTableRow).Value = yearlyChange ' prints in Column J
                            ws.Range("J" & 2 + summaryTableRow).NumberFormat = "0.00" ' formats Column J
                            ws.Range("K" & 2 + summaryTableRow).Value = percentChange '  prints in Column K
                            ws.Range("K" & 2 + summaryTableRow).NumberFormat = "0.00%" '  formats Column K
                            ws.Range("L" & 2 + summaryTableRow).Value = totalStockVolume '  prints in Column L (total stock volume)
                            ws.Range("L" & 2 + summaryTableRow).NumberFormat = "#,###" '  formats Column L
                            
                            ' color formatting for yearChange
                            If yearlyChange > 0 Then
                                ws.Range("J" & 2 + summaryTableRow).Interior.ColorIndex = 4 ' green for positive
                            ElseIf yearlyChange < 0 Then
                                ws.Range("J" & 2 + summaryTableRow).Interior.ColorIndex = 3 ' red for negative
                            Else
                                ws.Range("J" & 2 + summaryTableRow).Interior.ColorIndex = 0 ' white for no change
                            End If
                            
                        End If
                        
                        '   reset the values of the total stock value
                        totalStockVolume = 0
                        '   reset the values of yearly change
                        yearlyChange = 0
                        ' move to the next row in the summary table
                        summaryTableRow = summaryTableRow + 1
                    
                    '   check to see if the ticker is the same
                    Else
                        totalStockVolume = totalStockVolume + ws.Cells(row, 7).Value ' grabs stock volume from column G
                    End If
                
                Next row
                
        Next ws
          
End Sub

