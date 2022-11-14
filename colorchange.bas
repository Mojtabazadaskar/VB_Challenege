Attribute VB_Name = "Module1"
Sub ColorCahnge()

Summary_Table_row = 2


lastRow = Cells(Rows.Count, 1).End(xlUp).Row

For i = 2 To lastRow
   If Cells(i, 12) > 0 Then
            Cells(i, 12).Interior.ColorIndex = 4
            ElseIf Cells(i, 12) < 0 Then
            Cells(i, 12).Interior.ColorIndex = 3
            Else
           Cells(i, 12).Interior.ColorIndex = 2
            End If
  Next i

End Sub

