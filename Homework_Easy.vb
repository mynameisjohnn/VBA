Sub stocks()

  ' set initial variable for holding ticker names
  Dim ticker As String

  ' set initial variable for holding total volume per ticker name
  Dim volume As Double
  volume = 0

  ' keep track of the location for each ticker name in table
  Dim summary As Integer
  summary = 2

  ' define the last row
  Dim LastRow As Long
  LastRow = Cells(Rows.Count, 1).End(xlUp).Row

  ' loop through all tickers to the last row
  For i = 2 To LastRow

        'check if still within the same stock name, if it is not...
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    
          ' set the ticker name
          ticker = Cells(i, 1).Value
    
          ' add to the volume total
          volume = volume + Cells(i, 3).Value
    
          ' print the ticker in the summary
          Range("I" & summary).Value = ticker
    
          'print the volume total to the summary
          Range("J" & summary).Value = volume
    
          ' add one to the summary row
          summary = summary + 1
    
          ' reset the volume
          volume = 0
        
        ' if the cell immediately following a row is the same ticker...
        Else

          ' add to the volume total
          volume = volume + Cells(i, 3).Value

        End If
      
    Next i

End Sub
  
