Sub multiple_year_stock_data()

For Each ws In Worksheets
       
        Dim Work_sheet_Name As String
        
        Dim i, j, Tick_Count, Last_Row_CA, Last_Row_CI, Per_Change, GreatIncr, Great_Decr, Great_Vol As Double
                      
        Work_sheet_Name = ws.Name
                
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
       
        Tick_Count = 2
        
        j = 2
         
        Last_Row_CA = ws.Cells(Rows.Count, 1).End(xlUp).Row
            
            For i = 2 To Last_Row_CA
            
                If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                
                    ws.Cells(Tick_Count, 9).Value = ws.Cells(i, 1).Value
                
                    ws.Cells(Tick_Count, 10).Value = ws.Cells(i, 6).Value - ws.Cells(j, 3).Value
                
                If ws.Cells(Tick_Count, 10).Value > 0 Then
                
                    ws.Cells(Tick_Count, 10).Interior.ColorIndex = 4
                
                Else
                
                    ws.Cells(Tick_Count, 10).Interior.ColorIndex = 3
                
                End If
                    
                If ws.Cells(j, 3).Value <> 0 Then
                    
                    Per_Change = ((ws.Cells(i, 6).Value - ws.Cells(j, 3).Value) / ws.Cells(j, 3).Value)
                    
                    ws.Cells(Tick_Count, 11).Value = Format(Per_Change, "Percent")
                    
                Else
                    
                    ws.Cells(Tick_Count, 11).Value = Format(0, "Percent")
                    
                End If
                    
                    ws.Cells(Tick_Count, 12).Value = WorksheetFunction.Sum(Range(ws.Cells(j, 7), ws.Cells(i, 7)))
                
                    Tick_Count = Tick_Count + 1
                
                    j = i + 1
                
                End If
            
            Next i
            
            Last_Row_CI = ws.Cells(Rows.Count, 9).End(xlUp).Row
       
        Great_Vol = ws.Cells(2, 12).Value
        Great_Incr = ws.Cells(2, 11).Value
        Great_Decr = ws.Cells(2, 11).Value
        
            For i = 2 To Last_Row_CI
            
                If ws.Cells(i, 12).Value > Great_Vol Then
                Great_Vol = ws.Cells(i, 12).Value
                ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
                
            Else
                
                Great_Vol = Great_Vol
                
            End If
                
                If ws.Cells(i, 11).Value > Great_Incr Then
                Great_Incr = ws.Cells(i, 11).Value
                ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
                
                Else
                
                Great_Incr = Great_Incr
                
                End If
                
                If ws.Cells(i, 11).Value < Great_Decr Then
                Great_Decr = ws.Cells(i, 11).Value
                ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
                
                Else
                
                Great_Decr = Great_Decr
                
                End If
                
            ws.Cells(2, 17).Value = Format(Great_Incr, "Percent")
            ws.Cells(3, 17).Value = Format(Great_Decr, "Percent")
            ws.Cells(4, 17).Value = Format(Great_Vol, "Scientific")
            
            Next i
            
        Worksheets(Work_sheet_Name).Columns("A:Z").AutoFit
            
    Next ws
        
End Sub
