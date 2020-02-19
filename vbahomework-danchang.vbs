Sub vbahomework():

For Each ws In Worksheets

'Setting the headers
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"

'Define variables
    Dim lastRow As String
    Dim totalvol As Double
    Dim summaryrow As Integer
    Dim StartPrice As Double
    Dim lastprice As Double
    Dim greatincr As Double
    Dim greatdecr As Double
    Dim greatvol As Double

'Set variables
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row 'finding last row in the data set
    totalvol = 0 'initial value of total
    summaryrow = 2 'the row which we will first input data in the summary table
    StartPrice = ws.Cells(2, 3).Value 'this gives you the first start value -- setting it the first time
    greatincr = 0 'start base value of greatest increase
    greatdecr = 0 'start base value of greatest decrease
    greatvol = 0 'start base value of greatest volume

'Start for loop
        For i = 2 To lastRow
    
            If StartPrice <= 0 Then
                StartPrice = ws.Cells(i + 1, 3).Value
            End If

            totalvol = (totalvol + ws.Cells(i, 7).Value)
        
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                lastprice = ws.Cells(i, 6).Value
                ws.Cells(summaryrow, 9).Value = ws.Cells(i, 1).Value 'ticker
                ws.Cells(summaryrow, 10).Value = lastprice - StartPrice 'yearly change
                
                'Format yearly change to green if >0 and red if <0
                If ws.Cells(summaryrow, 10).Value > 0 Then
                    ws.Cells(summaryrow, 10).Interior.ColorIndex = 4
                Else: ws.Cells(summaryrow, 10).Interior.ColorIndex = 3
                End If
            
                ws.Cells(summaryrow, 11).Value = (lastprice / StartPrice) - 1 '%  change
                
                'Format % change column as percentage
                ws.Cells(summaryrow, 11).NumberFormat = "0.00%"
                
                'Determine which ticker has greatest increase and decrease
                If ws.Cells(summaryrow, 11).Value > greatincr Then
                    greatincr = ws.Cells(summaryrow, 11).Value
                    ws.Cells(2, 17).Value = greatincr
                    ws.Cells(2, 16).Value = ws.Cells(i, 1).Value
                    ws.Cells(2, 17).NumberFormat = "0.00%"
                ElseIf ws.Cells(summaryrow, 11).Value < greatdecr Then
                    greatdecr = ws.Cells(summaryrow, 11).Value
                    ws.Cells(3, 17).Value = greatdecr
                    ws.Cells(3, 16).Value = ws.Cells(i, 1).Value
                    ws.Cells(3, 17).NumberFormat = "0.00%"
                End If
                
                ws.Cells(summaryrow, 12).Value = totalvol 'total volume
                'Determine which ticker has the greatest volume
                If ws.Cells(summaryrow, 12).Value > greatvol Then
                    greatvol = ws.Cells(summaryrow, 12).Value
                    ws.Cells(4, 17).Value = greatvol
                    ws.Cells(4, 16).Value = ws.Cells(i, 1).Value
                    ws.Cells(4, 17).NumberFormat = "General"
                End If

                'Setting values for next iteration
                summaryrow = summaryrow + 1
                StartPrice = ws.Cells(i + 1, 3).Value 'this gives you the next start price after every iteration
                totalvol = 0

            End If
    
        Next i
'Format column width
    ws.Cells.SpecialCells(xlCellTypeVisible).EntireColumn.AutoFit

Next ws

End Sub


