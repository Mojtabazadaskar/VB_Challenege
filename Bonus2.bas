Attribute VB_Name = "Module3"
Sub Bonus2()

Dim MaxIncrease As Double
Dim MaxDecrease As Double
Dim MaxVol As Double



 


For i = 2 To 3500

 ' Max Increase
            MaxIncrease = WorksheetFunction.Max(Range("M:M"))
            Range("S10") = MaxIncrease
            Range("s10").NumberFormat = "000%"




    If (Cells(i, 13).Value = MaxIncrease) Then
       Range("R10") = Cells(i, 11).Value
   End If

'Max Decrease
            MaxDecrease = WorksheetFunction.Min(Range("M:M"))
            Range("S11") = MaxDecrease
            Range("S11").NumberFormat = "00%"


    If (Cells(i, 13).Value = MaxDecrease) Then
       Range("R11") = Cells(i, 11).Value
   End If
   
'Greatest Vol
        MaxVol = WorksheetFunction.Max(Range("N:N"))
        Range("S12") = MaxVol
        
               
   If (Cells(i, 13).Value = MaxVol) Then
       Range("R12") = Cells(i, 11).Value
   End If
   
Next i

End Sub
