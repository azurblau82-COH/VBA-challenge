Sub ticker()

' Set an initial variable for holding the ticker
  Dim ticker As String
Lastrow = Cells(Rows.Count, 1).End(xlUp).Row
  Dim tablerow As Integer
tablerow = 2
Cells(1, 8).Value = "Ticker"
  ' Loop through all credit card purchases
  For i = 2 To Lastrow

        ' Check if we are still within the same ticker, if we are not...
    
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

        'input ticker to last column
      
        ticker = Cells(i, 1).Value
        
        Range("H" & tablerow).Value = ticker
        tablerow = tablerow + 1
        End If
 Next i

End Sub