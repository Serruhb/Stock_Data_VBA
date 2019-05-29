#Easy VBA Script
```Sub easySolution()

    Dim lastRow As Long
    Dim currRow As Integer
    Dim totalVol As Double
       
    For Each ws In Worksheets
        'assigns headers to new information
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Total Volume"
          
        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        currRow = 2
        
        
        For i = 2 To lastRow
            
            totalVol = totalVol + ws.Cells(i, 7).Value
            'compare the tickers to find where each one ends
            If ws.Cells(i + 1, 1) <> ws.Cells(i, 1) Then
                
                'assigns values to new table info
                ws.Cells(currRow, 9).Value = ws.Cells(i, 1).Value
                ws.Cells(currRow, 10).Value = totalVol
                totalVol = 0
                currRow = currRow + 1
            End If
        Next i
    Next ws
End Sub```

#Moderate VBA Script

```Sub moderateSolution()



    Dim lastRow As Long
    Dim lastRow2 As Long
    Dim currRow As Integer
    Dim totalVol As Double
    Dim count As Double
    
    
        
        
 For Each ws In Worksheets
 
    
        'assigns headers to new information
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Volume"

        
        lastRow = ws.Cells(Rows.count, 1).End(xlUp).Row
        currRow = 2
        
        
        For i = 2 To lastRow
            
            totalVol = totalVol + ws.Cells(i, 7).Value
            count = count + 1
            'compare the tickers
            
            If ws.Cells(i + 1, 1) <> ws.Cells(i, 1) Then
                
                'assigns values to new table info
                ws.Cells(currRow, 9).Value = ws.Cells(i, 1).Value
                ws.Cells(currRow, 10).Value = ws.Cells(i, 6).Value - ws.Cells(i - count + 1, 3).Value
                ws.Cells(currRow, 12).Value = totalVol
    
                'added this for the error of dividing by zero
                If ws.Cells(i - count + 1, 3).Value = 0 Then
                    ws.Cells(currRow, 11).Value = 0
                Else
                    ws.Cells(currRow, 11).Value = ws.Cells(currRow, 10).Value / ws.Cells(i - count + 1, 3).Value
                End If
                
                'conditional if greater than 0 green, less than red
                If ws.Cells(currRow, 10).Value > 0 Then
                    ws.Cells(currRow, 10).Interior.Color = RGB(0, 255, 59)
                ElseIf ws.Cells(currRow, 10).Value < 0 Then
                    ws.Cells(currRow, 10).Interior.Color = RGB(255, 0, 0)
                End If
                
                totalVol = 0
                currRow = currRow + 1
                count = 0
                
            End If
        Next i
        lastRow2 = ws.Cells(Rows.count, 9).End(xlUp).Row
        ws.Range("K2:K" & lastRow2).NumberFormat = "0.00%"
    Next ws
End Sub```

#Hard VBA SCript

```Sub hardSolution()

    Dim lastRow As Long
    Dim lastRow2 As Long
    Dim currRow As Integer
    Dim totalVol As Double
    Dim count As Double
        
        
 For Each ws In Worksheets
 
    
        'assigns headers to new information
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Volume"
        ws.Cells(2, 14).Value = "Greatest Increase"
        ws.Cells(3, 14).Value = "Greatest Decrease"
        ws.Cells(4, 14).Value = "Greatest Total Volume"
        ws.Cells(1, 15).Value = "Ticker"
        ws.Cells(1, 16).Value = "Value"
        
        lastRow = ws.Cells(Rows.count, 1).End(xlUp).Row
        currRow = 2
        
        
        For i = 2 To lastRow
            
            totalVol = totalVol + ws.Cells(i, 7).Value
            count = count + 1
            'compare the tickers
            
            If ws.Cells(i + 1, 1) <> ws.Cells(i, 1) Then
                
                'assigns values to new table info
                ws.Cells(currRow, 9).Value = ws.Cells(i, 1).Value
                ws.Cells(currRow, 10).Value = ws.Cells(i, 6).Value - ws.Cells(i - count + 1, 3).Value
                ws.Cells(currRow, 12).Value = totalVol
    
                'added this for the error of dividing by zero
                If ws.Cells(i - count + 1, 3).Value = 0 Then
                    ws.Cells(currRow, 11).Value = 0
                Else
                    ws.Cells(currRow, 11).Value = ws.Cells(currRow, 10).Value / ws.Cells(i - count + 1, 3).Value
                End If
                
                'conditional if greater than 0 green, less than red
                If ws.Cells(currRow, 10).Value > 0 Then
                    ws.Cells(currRow, 10).Interior.Color = RGB(0, 255, 59)
                ElseIf ws.Cells(currRow, 10).Value < 0 Then
                    ws.Cells(currRow, 10).Interior.Color = RGB(255, 0, 0)
                End If
                
                totalVol = 0
                currRow = currRow + 1
                count = 0
                
            End If
        Next i
        
        lastRow2 = ws.Cells(Rows.count, 9).End(xlUp).Row
        ws.Range("K2:K" & lastRow2).NumberFormat = "0.00%"
        ws.Cells(2, 16).NumberFormat = "0.00%"
        ws.Cells(3, 16).NumberFormat = "0.00%"
    
        For i = 2 To lastRow2
    'return the greatest inc %
                If ws.Cells(2, 16).Value < ws.Cells(i, 11).Value Then
                    ws.Cells(2, 16).Value = ws.Cells(i, 11).Value
                    ws.Cells(2, 15).Value = ws.Cells(i, 9).Value
                End If
    'return the greatest dec %
                If ws.Cells(3, 16).Value > ws.Cells(i, 11).Value Then
                    ws.Cells(3, 16).Value = ws.Cells(i, 11).Value
                    ws.Cells(3, 15).Value = ws.Cells(i, 9).Value
                End If
                
    'returns the greatest total amount
                If ws.Cells(4, 16).Value < ws.Cells(i, 12).Value Then
                    ws.Cells(4, 16).Value = ws.Cells(i, 12).Value
                    ws.Cells(4, 15).Value = ws.Cells(i, 9).Value
                End If
        Next i
        
    Next ws

End Sub```
