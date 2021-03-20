Attribute VB_Name = "Module11"
Sub StockSolution()
'Homework
    ' Create a script that will loop through all the stocks for one year and output
        ' 1. the ticker symbol
        ' 2. Yearly change from opening price at the beginning of a given year to the closing
        'price at the end of that year
        '3. The total stock volume of the stock
        '4. Conditional formatting to highight positive change in green and neg in red
        
'Add ws. to add headers and data on each worksheet
  Dim ws As Worksheet
  For Each ws In Worksheets
  
'Set the headers on the new columns
  ws.Range("I1").Value = "Ticker"
  ws.Range("I1").Font.Bold = True
  ws.Range("J1").Value = "Yearly Change"
  ws.Range("J1").Font.Bold = True
  ws.Range("K1").Value = "Percent Change"
  ws.Range("K1").Font.Bold = True
      'Get the decimal and percent format
  ws.Columns("J").NumberFormat = "0.00"
  ws.Columns("K").NumberFormat = "0.00%"
  
  ws.Range("L1").Value = "Total Stock Volume"
  ws.Range("L1").Font.Bold = True
  ws.Range("I:L").Columns.AutoFit
  
'Set last row
  LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
' Set the initial variables for each item: ticker, total, year change, percent change
  Dim Ticker_Name As String
  Dim Stock_Total As LongLong
    Stock_Total = 0
  Dim Yearly_Change As Double

  Dim Percent_Change As Double
        
  Dim Yearopen As Double
    Yearopen = ws.Cells(2, 3).Value
       
  Dim Yearclose As Double
  
  'Location of ticker data in the summary table
  Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2
     
  ' Loop through all ticker symbols
  For r = 2 To LastRow

  ' Check if the ticker symbol in the next row is not the same
  If ws.Cells(r + 1, 1).Value <> ws.Cells(r, 1).Value Then

      ' Set the ticker name
      Ticker_Name = ws.Cells(r, 1).Value

      ' Add to the Stock Total
      Stock_Total = Stock_Total + ws.Cells(r, 7).Value
      
      'Get the yearly change per ticker symbol
        'Calculate yearly change and percent change
      Yearclose = ws.Cells(r, 6).Value
      Yearly_Change = (Yearclose - Yearopen)
      Percent_Change = (Yearly_Change / Yearopen)
            
             
      ' Print the ticker name in the Summary Table
      ws.Range("I" & Summary_Table_Row).Value = Ticker_Name

      ' Print the stock total to the Summary Table
      ws.Range("L" & Summary_Table_Row).Value = Stock_Total
      
      'Print the Yearly Change to the Summary Table
      ws.Range("J" & Summary_Table_Row).Value = Yearly_Change
         
          
      'Print the Percent Change to the Summary Table
      ws.Range("K" & Summary_Table_Row).Value = Percent_Change
      
      'Add conditional formatting to yearly change column
    
     If ws.Range("J" & Summary_Table_Row).Value > 0 Then
        ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
     ElseIf ws.Range("J" & Summary_Table_Row).Value < 0 Then
        ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
      End If
      
       'Update yearopen for next ticker
      Yearopen = ws.Cells(r + 1, 3).Value

      ' Add one to the summary table row
      Summary_Table_Row = Summary_Table_Row + 1
      
      
      ' Reset the Stock Total
      Stock_Total = 0
      
      
    ' If the ticker is the same
    Else

      ' Add to the Stock Total
      Stock_Total = Stock_Total + ws.Cells(r, 7).Value

      
    End If

    Next r
      
  Next ws

End Sub

