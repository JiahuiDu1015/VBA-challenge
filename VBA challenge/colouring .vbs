Sub colouring()

  lastRow = ws.Cells(ws.Rows.Count, "K").End(xlUp).Row

  For i to lastRow

  If Colums("K").Value >= 0 Then

     Cells(i, 11).Interior.ColorIndex = 4

  ElseIf Colums("K").Value < 0 Then

     Cells(i, 11).Interior.ColorIndex = 4

  End If
  Next i

End Sub
