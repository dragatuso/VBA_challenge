Sub tickers()
    
    'Loop through all sheets
    For Each ws In Worksheets
        
    
        'Set the name of the columns and rows from the tables we are going to create
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
    
        'Define the variables
        Dim column As Long
        column = 1
    
        Dim tickercount As Long
        tickercount = 2
    
        Dim openpriceyear As Long
        openpriceyear = 2
    
        Dim greatestincrease As Double
        greatestincrese = 0
    
        Dim tickernamegreatestincrease As String
    
        Dim greatestdecrease As Double
        greatestdecrease = 0
    
        Dim tickernamegreatestdecrease As String
    
        Dim greatestvolume As Double
        greatestvolume = 0
    
        Dim tickernamegreatestvolume As String
      

        'Define the last row of the worksheet
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
        'Create the rows of the first additional table we are creating
        For i = 2 To lastrow
    
            ws.Cells(tickercount, 10).Value = ws.Cells(i, 6).Value - ws.Cells(openpriceyear, 3).Value
            
                
            ws.Cells(tickercount, 11).Value = (ws.Cells(i, 6).Value / ws.Cells(openpriceyear, 3).Value) - 1
        
        
            ws.Cells(tickercount, 12).Value = ws.Cells(tickercount, 12) + ws.Cells(i, 7).Value
        
                          
                
            If ws.Cells(i + 1, column).Value <> ws.Cells(i, column).Value Then
                ws.Cells(tickercount, 9).Value = ws.Cells(i, column).Value
            
                'Yearly change format
                If ws.Cells(tickercount, 10).Value < 0 Then
                    ws.Range("J" & tickercount).Interior.ColorIndex = 3
                ElseIf ws.Cells(tickercount, 10).Value > 0 Then
                    ws.Range("J" & tickercount).Interior.ColorIndex = 4
                Else
                    ws.Range("J" & tickercount).Interior.ColorIndex = 0
                
                End If
        
                'Percentage Change format
                ws.Range("K" & tickercount).NumberFormat = "0.00%"
                If ws.Cells(tickercount, 11).Value < 0 Then
                    ws.Range("K" & tickercount).Interior.ColorIndex = 3
                ElseIf ws.Cells(tickercount, 11).Value > 0 Then
                    ws.Range("K" & tickercount).Interior.ColorIndex = 4
                Else
                    ws.Range("K" & tickercount).Interior.ColorIndex = 0
                
                End If
                                    
                     
                tickercount = tickercount + 1
                openpriceyear = i + 1
                
                                    
            End If
        
        
        
        Next i
    
       
        'Calculate the values of the cells in the additional table we are creating
        For i = 2 To tickercount
        
            If ws.Cells(i, 11).Value > greatestincrease Then
                greatestincrease = ws.Cells(i, 11).Value
                tickernamegreatestincrease = ws.Cells(i, 9).Value
            
            End If
            
            If ws.Cells(i, 11).Value < greatestdecrease Then
                greatestdecrease = ws.Cells(i, 11).Value
                tickernamegreatestdecrease = ws.Cells(i, 9).Value
            
            End If
            
            If ws.Cells(i, 12).Value > greatestvolume Then
                greatestvolume = ws.Cells(i, 12).Value
                tickernamegreatestvolume = ws.Cells(i, 9).Value
        
            End If
        
        Next i
    
        'Input the calculated values in the new table and format them
        ws.Range("Q2").Value = greatestincrease
        ws.Range("Q2").NumberFormat = "0.00%"
        ws.Range("P2").Value = tickernamegreatestincrease
        ws.Range("Q3").Value = greatestdecrease
        ws.Range("Q3").NumberFormat = "0.00%"
        ws.Range("P3").Value = tickernamegreatestdecrease
        ws.Range("Q4").Value = greatestvolume
        ws.Range("P4").Value = tickernamegreatestvolume
    
    Next ws

End Sub