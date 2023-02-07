Attribute VB_Name = "Module11"
Sub tickers():

    For Each ws In Worksheets ' Repeats entire Sub on each worksheet

        ws.Range("I1").Value = "Ticker"                 '
        ws.Range("J1").Value = "Yearly Change"            '
        ws.Range("K1").Value = "Percent Change"             '
        ws.Range("L1").Value = "Total Stock Volume"           '
        ws.Range("O2").Value = "Greatest % Increase"            ' Labeling our rows
        ws.Range("O3").Value = "Greatest % Decrease"          '
        ws.Range("O4").Value = "Greatest Total Volume"      '
        ws.Range("P1").Value = "Ticker"                   '
        ws.Range("Q1").Value = "Value"                  '
    
        ws.Range("A:Q").Columns.AutoFit
    
        lastRow = ws.Cells(Rows.count, 1).End(xlUp).row ' Finds last row of dataset
        
        Dim tickerName As String
        Dim yrChange As Double
        Dim totalVol As Double
        totalVol = 0
        Dim rateOpen As Double
        Dim rowOpen As Double
        rowOpen = 2
        rateOpen = ws.Cells(2, 3).Value
        Dim rateClose As Double
        Dim tickerRows As Double
        tickerRows = 2
        Dim row As Long
        Dim perChange As Double
        
        
        For row = 2 To lastRow
        
        
        
            If ws.Cells(row + 1, 1).Value <> ws.Cells(row, 1).Value Then
                
                tickerName = ws.Cells(row, 1).Value
                totalVol = totalVol + ws.Cells(row, 7).Value
            
                ws.Cells(tickerRows, 9).Value = tickerName
                ws.Cells(tickerRows, 12).Value = totalVol
                totalVol = 0
            
                rateClose = ws.Cells(row, 6).Value ' Defines the close rate before change
                yrChange = rateClose - rateOpen ' Determines yearly change
                ws.Cells(tickerRows, 10).Value = yrChange
                    
                    If yrChange < 0 Then
                        ws.Cells(tickerRows, 10).Interior.ColorIndex = 3 ' Changes color of negative yrChange cells to red
                    Else
                        ws.Cells(tickerRows, 10).Interior.ColorIndex = 4 ' Changes color of positive yrChange to green
                    End If
                    
                perChange = ((rateClose - rateOpen) / rateOpen) ' Calculates percent change
                ws.Cells(tickerRows, 11).Value = perChange
                ws.Cells(tickerRows, 11).NumberFormat = "0.00%"
                    
                rateOpen = ws.Cells(row + 1, 3).Value
                
                tickerRows = tickerRows + 1
            Else
                totalVol = totalVol + ws.Cells(row, 7).Value
            End If
            
            If ws.Cells(rowOpen, 3).Value = 0 Then ' Checks where rowOpen starts, skips 0-value rowOpen
                For rowOpen = rowOpen To lastRow
                    If ws.Cells(rowOpen, 3).Value <> 0 Then
                        rowOpen = rowOpen + 1
                        Exit For
                    
                    End If
            
                Next rowOpen
            
            End If
            
        Next row
        
        lastNew = ws.Cells(Rows.count, 9).End(xlUp).row 'Finds last row of output data's rows
        
        Dim grIncrease As Double
        grIncrease = 0
        Dim grDecrease As Double
        grDecrease = 0
        Dim grTotVol As Double
        grTotVol = 0
        
        Dim percentRow As Integer
        percentRow = 2
        
        For row = 2 To lastNew
        
            If ws.Cells(row, 11) > grIncrease Then ' Calculates the Greatest Percent Increase
                grIncrease = ws.Cells(row, 11).Value
                ws.Range("Q2").Value = grIncrease
                ws.Range("Q2").NumberFormat = "0.00%"
            End If
            
            If ws.Cells(row, 11) < grDecrease Then ' Calculates the Greatest Percent Decrease
                grDecrease = ws.Cells(row, 11).Value
                ws.Range("Q3").Value = grDecrease
                ws.Range("Q3").NumberFormat = "0.00%"
            End If

            If ws.Cells(row, 12) > grTotVol Then ' Calculates the Greatest Total Volume
                grTotVol = ws.Cells(row, 12).Value
                ws.Range("Q4").Value = grTotVol
                ws.Range("Q4").NumberFormat = "0"
            End If
            
            percentRow = percentRow + 1
            
        Next row
        
        Dim tickerIncrease As Double
        tickerIncrease = Application.Match(ws.Range("Q2").Value, ws.Range("K2:K3001"), 0) 'Adds the ticker labels next to the caculated greatest percent increase, decrease, and total
        ws.Range("P2").Value = ws.Range("I" & tickerIncrease + 1)
        
        Dim tickerDecrease As Double
        tickerDecrease = Application.Match(ws.Range("Q3").Value, ws.Range("K2:K3001"), 0)
        ws.Range("P3").Value = ws.Range("I" & tickerDecrease + 1)
        
        Dim tickerTotal As Double
        tickerTotal = Application.Match(ws.Range("Q4").Value, ws.Range("L2:L3001"), 0)
        ws.Range("P4").Value = ws.Range("I" & tickerTotal + 1)
        
    Next ws

End Sub
