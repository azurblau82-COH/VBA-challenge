Sub yearly_change()
'define variables
Dim open_price As Double
Dim close_price As Double

'set open_price to C2
open_price = Cells(2, 3).Value


Dim yearlychange As Double
Dim percentchange As Double
Dim stockvolume As Double


Lastrow = Cells(Rows.Count, 1).End(xlUp).Row

  Dim tablerow As Integer
tablerow = 2
Cells(1, 9).Value = "Yearly Change"
Cells(1, 10).Value = "Percent Change"
Cells(1, 11).Value = "Stock Volume"
stockvolume = 0

  For i = 2 To Lastrow
  
  Cells(tablerow, 10).NumberFormat = "0.00"

    stockvolume = stockvolume + Cells(i, 7).Value
    
           
        ' Check if we are still within the same ticker, if we are not...
    
    
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        
            close_price = Cells(i, 6).Value
            yearlychange = close_price - open_price
            percentchange = (yearlychange / open_price) * 100
            
            
            Cells(tablerow, 9).Value = yearlychange
            Cells(tablerow, 10).Value = percentchange
            Cells(tablerow, 11).Value = stockvolume
            
                If yearlychange < 0 Then
                Cells(tablerow, 9).Interior.ColorIndex = 3
                Else
                Cells(tablerow, 9).Interior.ColorIndex = 4
                
                End If
                
            
            stockvolume = 0
            tablerow = tablerow + 1
            open_price = Cells(i + 1, 3).Value
            
            
            
        End If
    Next i
    
    

End Sub




