Sub YearlyChange()

For Each ws In Worksheets

  Dim Row As Long
  ' Set an initial variable for holding the tickersymbol
  Dim TickerSymbol As String
  ' Set an initial variable for holding the volume
 
 ' Keep track of the location for each credit card brand in the summary table
  Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2
  

  
'set last row variable
lr = ws.Range("A2").End(xlDown).Row

' Loop through all TickerSymbols
  For Row = 2 To lr
         
  'Check if we are still within the same TickerSymbol, if not...
    If ws.Cells(Row + 1, 1).Value <> ws.Cells(Row, 1).Value Then
   
   
      ' Set the TickerSymbol
      ws.Cells(1, 9).Value = "Ticker"
      TickerSymbol = ws.Cells(Row, 1).Value
    ' Print the TickerSymbol in the Summary Table
      ws.Range("I" & Summary_Table_Row).Value = TickerSymbol

     ' Add to the VolumeTotal
      VolumeTotal = VolumeTotal + ws.Cells(Row, 7).Value
      ' Print the VolumeAmount to the Summary Table
      ws.Cells(1, 12).Value = "Total Stock Volume"
      ws.Range("L" & Summary_Table_Row).Value = VolumeTotal
  
      ' Add one to the summary table row
      Summary_Table_Row = Summary_Table_Row + 1
      
      'Reset the TotalVolume
       VolumeTotal = 0
    
' If the cell immediately following a row is the same Ticker
    Else

      ' Add to the Volume Total
      VolumeTotal = VolumeTotal + ws.Cells(Row, 7).Value
     
     End If

  Next Row
Next ws
End Sub

Sub YearlyPriceChange()

For Each ws In Worksheets

' Keep track of the location for each Ticker in the summary table
Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2
Dim Row As Long
Dim Counter As Long
Counter = 0
Dim OpenPrice As Double
Dim ClosePrice As Double


'set last row variable
lr = ws.Range("A2").End(xlDown).Row

' Loop through all TickerSymbols
  For Row = 2 To lr
  
 'Check if we are still within the same TickerSymbol & assign Close & Open Price variables
    If ws.Cells(Row + 1, 1).Value <> ws.Cells(Row, 1).Value Then
    ClosePrice = ws.Cells(Row, 6).Value
    OpenPrice = ws.Cells(Row - Counter, 3).Value
    
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Range("J" & Summary_Table_Row) = ClosePrice - OpenPrice
    If ws.Range("J" & Summary_Table_Row).Value > 0 Then
    ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
    Else: ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
   
    End If
    If ws.Range("J" & Summary_Table_Row) = 0 Then
    ws.Range("K" & Summary_Table_Row) = 0
    Else
    ws.Cells(1, 11).Value = "Percentage Change"
    ws.Range("K" & Summary_Table_Row) = ((ClosePrice - OpenPrice) / OpenPrice)
    ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
    End If
 ' Add one to the summary table row
      Summary_Table_Row = Summary_Table_Row + 1
        
    Counter = 0
      
    Else
    
    Counter = Counter + 1
    
End If

 Next Row
Next ws
End Sub

Sub GreatestOnTab()

For Each ws In Worksheets
'hardcoding table
ws.Range("O2") = "Greatest % Increase"
ws.Range("O3") = "Greatest % Decrease"
ws.Range("O4") = "Greatest Total Volume"
ws.Range("P1") = "Ticker"
ws.Range("Q1") = "Value"

'Getting Greatest % Increase and Ticker
Dim StartingPercentage As Double
StartingPercentage = ws.Range("K2").Value

'set last row variable
lr = ws.Range("I2").End(xlDown).Row

'setting variables for nested for lop
Dim Row As Integer

For Row = 2 To lr
   If ws.Cells(Row, 11) > StartingPercentage Then
   StartingPercentage = ws.Cells(Row, 11).Value
   ws.Range("Q2") = ws.Cells(Row, 11).Value
   ws.Range("Q2").NumberFormat = "0.00%"
   ws.Range("P2") = ws.Cells(Row, 9).Value
   End If
Next Row
Next ws
End Sub

Sub LeastOnTab()

For Each ws In Worksheets

'Getting Greatest % Decrease and Ticker
Dim StartingPercentage As Double
StartingPercentage = ws.Range("K2").Value

'set last row variable
lr = ws.Range("I2").End(xlDown).Row

'setting variables for nested for lop
Dim Row As Integer

For Row = 2 To lr
   If ws.Cells(Row, 11) < StartingPercentage Then
   StartingPercentage = ws.Cells(Row, 11).Value
   ws.Range("Q3") = ws.Cells(Row, 11).Value
   ws.Range("Q3").NumberFormat = "0.00%"
   ws.Range("P3") = ws.Cells(Row, 9).Value
   End If
Next Row
Next ws
End Sub
Sub MaxVolume()

For Each ws In Worksheets

'Getting Greatest Total Volume and Ticker
Dim StartingVolume As Double
StartingVolume = ws.Range("K2").Value

'set last row variable
lr = ws.Range("I2").End(xlDown).Row

'setting variables for nested for lop
Dim Row As Integer

For Row = 2 To lr
   If ws.Cells(Row, 12) > StartingVolume Then
   StartingVolume = ws.Cells(Row, 12).Value
   ws.Range("Q4") = ws.Cells(Row, 12).Value
   ws.Range("P4") = ws.Cells(Row, 9).Value
   End If
Next Row
Next ws
End Sub

