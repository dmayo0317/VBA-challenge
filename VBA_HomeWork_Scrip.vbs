Sub Yearly_Stock()

'Loops throw each of the worksheets in the workbook'
  For Each ws In Worksheets
  
  'Makes the worksheet active'
  ws.Activate
  
  'sets the header names for each worksheets'
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
     ws.Range("L1").Value = "Total Stock Volume"
        
    'sets the Format of Colum K as a Percent "%" '
    ws.Range("K:K").NumberFormat = "0.00%"
      
  ' Setting up each indivdual varibales'
  Dim ticker_Name As String
  Dim ticker_Total As Double
  Dim opening_price As Double
  Dim closing_price As Double
  Dim DeltaChange As Double
  Dim DeltaChangePercent As Double
    
  'seting up the varaibles to hold a value at 0 First'
  opening_price = 0
  closing_price = 0
  ticker_volTotal = 0
  DeltaChange = 0
  
  ' Keep track of the location for each ticker name in the summary table
  
  Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2
 
  'Finds the last row of each worksheet'
    lastrow = ws.Cells(Rows.Count, "A").End(xlUp).Row
    
  For i = 2 To lastrow
  currentRow = Cells(i, 1).Value
  nextRow = Cells(i + 1, 1).Value
  
    'Start holdig the first opening in the data
  
    
    ' Check if we are still within the same credit card brand, if it is not...
    If nextRow <> currentRow Then
    
      ' Set the Ticker name, opening price and closing price'
      ticker_Name = Cells(i, 1).Value
      opening_price = Cells(i, 3).Value
      closing_price = Cells(i, 6).Value
      
      'calculate the delta between opening and closing'
      DeltaChange = closing_price - opening_price
      
      'Calculate the percet change between opening and closing while checking if there a 0 in the opening.
      If opening_price = 0 Then
        DeltaChangePercent = 0
        Else
        
        DeltaChangePercent = DeltaChange / opening_price
        
        'Hold the percent Change in the summary table'
            Range("K" & Summary_Table_Row).Value = DeltaChangePercent
        End If
              
      'Hold the percent Change in the summary table'
      Range("K" & Summary_Table_Row).Value = DeltaChangePercent
       
    'Change the color of the cell based on the percent (forloop Condition Check)'
       If DeltaChangePercent < 0 Then
              Range("K" & Summary_Table_Row).Interior.ColorIndex = 3
       Else
             Range("K" & Summary_Table_Row).Interior.ColorIndex = 4
      
      End If
               
      'Hold the delta change in the summary table'
      Range("J" & Summary_Table_Row).Value = DeltaChange

      'Add to the volume Total'
      ticker_volTotal = ticker_volTotal + Cells(i, 7).Value

      'Print the Ticker Name Summary in the summary Table'
      Range("I" & Summary_Table_Row).Value = ticker_Name

      'Print the Total Volume Amount to the Summary Table'
      Range("L" & Summary_Table_Row).Value = ticker_volTotal

      'Add one to the summary table row'
      Summary_Table_Row = Summary_Table_Row + 1
      
      'Reset the Variables to 0 '
      ticker_volTotal = 0
      opening_price = 0
      closing_price = 0
      DeltaChange = 0
      
    ' If the next sell is the same as the current sell then....'
    Else

      ' Add to the Brand Total
      ticker_volTotal = ticker_volTotal + Cells(i, 7).Value

    End If

  Next i
  
  Next ws


End Sub

