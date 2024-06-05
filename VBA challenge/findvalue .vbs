Sub FindValue()

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim greatestValue As Double
    Dim greatestName As String
    Dim greatestincrease As Double
    Dim greatestdecrease As Double
    Dim greatesttotalValue As String
    Dim currentNameincrease As String
    Dim currentNamedecrease As String
    Dim currentNamegreastvalue As String
    
    
    lastRow = ws.Cells(ws.Rows.Count, "M").End(xlUp).Row
    lastRow = ws.Cells(ws.Rows.Count, "N").End(xlUp).Row
   
    Set DataRange = ThisWorkbook.Sheets("Sheet1").Range("K2:K" & lastRow)
    Set DataRange = ThisWorkbook.Sheets("Sheet1").Range("M2:M" & lastRow)
    Set DataRange = ThisWorkbook.Sheets("Sheet1").Range("N2:N" & lastRow)
    
    greatestValue = -9999999
    smallestValue = 9999999
    

    
    For i = 2 To lastRow 
        currentNameincrease = ws.Cells(i, 18).Value
        currentNamedecrease = ws.Cells(i, 18).Value
        currentNamedecrease = ws.Cells(i, 18).Value
        greastestincrease = ws.Cells(i, 19).Value
        greastestdecrease = ws.Cells(i, 19).Value
        greatesttotalvalue = ws.Cells(i, 19).Value
        
       
        
        If currentNameincrease > greatestValue Then
            greatestValue = greastestincrease
            greatestName = currentNameincrease
            
        ElseIf currentNamedecrease < greatestValue Then
            greatestValue = greastestdecrease
            greatestName = currentNamedecrease
            
        Else 
        
        End If
        
    Next i
    
End Sub