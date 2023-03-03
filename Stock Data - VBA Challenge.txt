Sub MultipleYearStockData():
    
    For Each ws In Worksheets
    
        Dim WorksheetName As String
        Dim i As Long
        Dim j As Long
        Dim summaryrow As Long
        Dim LastRowA As Long
        Dim LastRowI As Long
        Dim PerChange As Double
        Dim GreatIncr As Double
        Dim GreatDecr As Double
        Dim GreatVol As Double
        

        WorksheetName = ws.Name

        ws.Range("I1").Value = "Thicker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percentage change"
        ws.Range("L1").Value = "Total Stock Volume"
        ws.Range("N2").Value = "Greatest % Increase"
        ws.Range("N3").Value = "Greatest % Decrease"
        ws.Range("N4").Value = "Greatest Total Volume"
        ws.Range("O1").Value = "Thicker"
        ws.Range("P1").Value = "Value"

        
        summaryrow = 2
        
        j = 2
        
        
        LastRowA = ws.Cells(Rows.Count, 1).End(xlUp).Row
      
        
            For i = 2 To LastRowA
            
                If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                
                ws.Cells(summaryrow, 9).Value = ws.Cells(i, 1).Value
                
                ws.Cells(summaryrow, 10).Value = ws.Cells(i, 6).Value - ws.Cells(j, 3).Value
               
                    If ws.Cells(summaryrow, 10).Value < 0 Then
                        ws.Cells(summaryrow, 10).Interior.ColorIndex = 3
                    Else
                        ws.Cells(summaryrow, 10).Interior.ColorIndex = 4
                    End If
                  
                  
                    If ws.Cells(j, 3).Value <> 0 Then
                        PerChange = ((ws.Cells(i, 6).Value - ws.Cells(j, 3).Value) / ws.Cells(j, 3).Value)
                        ws.Cells(summaryrow, 11).Value = Format(PerChange, "Percent")
                    Else
                        ws.Cells(summaryrow, 11).Value = Format(0, "Percent")
                    End If
             
                ws.Cells(summaryrow, 12).Value = WorksheetFunction.Sum(Range(ws.Cells(j, 7), ws.Cells(i, 7)))
            
                summaryrow = summaryrow + 1
                
                j = i + 1
                
                End If
            
            Next i
            
      
        LastRowI = ws.Cells(Rows.Count, 9).End(xlUp).Row
                
        GreatVol = ws.Cells(2, 12).Value
        GreatIncr = ws.Cells(2, 11).Value
        GreatDecr = ws.Cells(2, 11).Value
        
         
            For i = 2 To LastRowI
            
                If ws.Cells(i, 12).Value > GreatVol Then
                    GreatVol = ws.Cells(i, 12).Value
                    ws.Cells(4, 15).Value = ws.Cells(i, 9).Value
                Else
                    GreatVol = GreatVol
                End If
                
                
                If ws.Cells(i, 11).Value > GreatIncr Then
                    GreatIncr = ws.Cells(i, 11).Value
                    ws.Cells(2, 15).Value = ws.Cells(i, 9).Value
                Else
                    GreatIncr = GreatIncr
                
                End If
                
                If ws.Cells(i, 11).Value < GreatDecr Then
                    GreatDecr = ws.Cells(i, 11).Value
                    ws.Cells(3, 15).Value = ws.Cells(i, 9).Value
                Else
                    GreatDecr = GreatDecr
                End If
                
           
            ws.Range("P2").Value = Format(GreatIncr, "Percent")
            ws.Range("P3").Value = Format(GreatDecr, "Percent")
            ws.Range("P4").Value = Format(GreatVol, "Scientific")
            
            Next i
            
       
        Worksheets(WorksheetName).Columns("A:Z").AutoFit
            
    Next ws
        
End Sub