Sub vba_challenge()

Dim xSh As Worksheet
    Application.ScreenUpdating = False
    For Each xSh In Worksheets
        
        xSh.Select
   
 For Each ws In Worksheets
  lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
  
  ' set all my variables
  Dim ticker As String
  Dim opening_price As Single
   opening_price = 0
  Dim closing_price As Single
  Dim yearly_change As Single
  Dim percent_change As Double
  Dim total_stock_volume As Variant
  total_stock_volume = 0
  Dim Summary_Table_Row As Integer
   Summary_Table_Row = 2
   
   'make the headers of my table and get it ready
   Cells(1, 9).Value = "ticker"
   Cells(1, 10).Value = "yearly change"
   Cells(1, 11).Value = "percentage change"
   Cells(1, 12).Value = "total stock volume"
    Range("K2:K" & lastrow).NumberFormat = "0.00%"
    Range("j2:j" & lastrow).NumberFormat = "0.00"
   Columns("I:L").AutoFit

   ' Loop through all tickers info
   For i = 2 To lastrow
     
     If opening_price = 0 Then
     opening_price = Cells(i, 3).Value
    
     End If
     
      ' Check if we are still within the same ticker or not
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

       ' Set ticket value
          ticker = Cells(i, 1).Value
              
        'set what would be the closing price
        closing_price = Cells(i, 6).Value
        
       ' Add to total stock volume
         total_stock_volume = total_stock_volume + Cells(i, 7).Value
         
       'calculate the yearly change
       yearly_change = closing_price - opening_price
       
       'calculate the percentage change
       percent_change = yearly_change / opening_price
       
       ' Print all the info in the Summary Table
        Range("I" & Summary_Table_Row).Value = ticker
        Range("j" & Summary_Table_Row).Value = yearly_change
        Range("k" & Summary_Table_Row).Value = percent_change
        Range("L" & Summary_Table_Row).Value = total_stock_volume
        
        'conditional formating
         If yearly_change >= 0 Then
        
         Cells(Summary_Table_Row, 10).Interior.ColorIndex = 4
           Else
         Cells(Summary_Table_Row, 10).Interior.ColorIndex = 3
                          
        End If

         Summary_Table_Row = Summary_Table_Row + 1
         
                  
      ' Reset the total stock volume
         total_stock_volume = 0
         opening_price = 0
          
          
         ' If the cell immediately following a row is the same ticker..
         Else

         ' Add to the total stock volume
         total_stock_volume = total_stock_volume + Cells(i, 7).Value

        End If
        
        Next i
 Next ws
 Next

Application.ScreenUpdating = True

End Sub

