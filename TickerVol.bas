Attribute VB_Name = "Module2"

              
                    
  
    Sub TickerVOL()


Dim Total_Vol_Stock As Double
Dim Ticker_Symbol As String

Dim Summary_Next_Row As Integer

Dim OpenPrice As Double
Dim ClosePrice As Double

Dim MaxIncrease As Double
Dim MaxDecrease As Double
Dim MaxVol As Double

OpenPrice = Cells(2, 3).Value


Summary_Next_Row = 2
lastRow = Cells(Rows.Count, 1).End(xlUp).Row

' For loop to take out the Ticker name
For i = 2 To lastRow

         
         ClosePrice = Cells(i, 6).Value
                    
                    
  
              

    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
           ' Replacing the Values
         
            Ticker_Symbol = Cells(i, 1).Value
            ' getting the Total valuesss
            Total_Vol_Stock = Total_Vol_Stock + Cells(i, 7).Value
            
          
    'Adding the name and the total to the cells
            Range("K" & Summary_Next_Row).Value = Ticker_Symbol
            Range("N" & Summary_Next_Row).Value = Total_Vol_Stock
            Range("l" & Summary_Next_Row).Value = ClosePrice - OpenPrice
            Range("m" & Summary_Next_Row).Value = ClosePrice / OpenPrice - 1
            
            
            
    ' Asking to go for the next Row
             Summary_Next_Row = Summary_Next_Row + 1
             OpenPrice = Cells(i + 1, 3).Value
    ' changeing the Total for new ticker to Zero
           Total_Vol_Stock = 0
    ' if they Ticker are equal with the next one give us the name and the total
       Else
        Total_Vol_Stock = Total_Vol_Stock + Cells(i, 7).Value
    
    End If
   Cells(i, 13).NumberFormat = "0.00%"
  
   
    'cheking for the yearlychanges
        


Next i


   

   
   
    

End Sub

