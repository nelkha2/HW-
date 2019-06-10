Sub Ticker_Volumes_Moderate()

'Goal: Print per ticker: yearly change in opening price (color code), yearly change in opening price percentage, total volume


Dim ticker As String
Dim ticker_row_volume As Long
Dim ticker_row_price As Long
Dim total_volume As Double
Dim total_rows As Long
Dim Start_price As Double 'beginning of year price
Dim End_price As Double 'end of year price
Dim year_start As Long
Dim year_end As Long
Dim year_rows As Integer


'Column headers for data
Range("I1").Value = "Ticker"
Range("J1").Value = "Yearly Change"
Range("K1").Value = "Percentage Yearly Change"
Range("L1").Value = "Total Stock Volume"

'Calculating total rows containing data
Range("A2").Select
total_rows = ActiveCell.End(xlDown).Row

'starting points
total_volume = 0
ticker_row_volume = 2
ticker_row_price = 2
year_end = Application.Max(Range("b2:b" & total_rows + 1))
year_start = Application.Min(Range("b2:b" & total_rows + 1))


'Table structure for tickers and aggregate volume
For i = 2 To total_rows

  If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
  'identifying data
  ticker = Cells(i, 1).Value
  total_volume = total_volume + Cells(i, 7).Value


  'Printing data
  Range("I" & ticker_row_volume).Value = ticker
  Range("L" & ticker_row_volume).Value = total_volume

  'updating the Printing variables to be used in the next iteration
  ticker_row_volume = ticker_row_volume + 1
  total_volume = 0

  'aggregating volume when tickers are repeated
  Else
  total_volume = total_volume + Cells(i, 7).Value

  End If

Next i

'Output yearly change & its percentage
For i = 2 To total_rows

    If Cells(i, 1).Value = Cells(i + 1, 1).Value And Cells(i, 2).Value = year_start Then
    Start_price = Cells(i, 3).Value
    
    ElseIf Cells(i, 1).Value = Cells(i + 1, 1).Value And Cells(i, 2).Value <> year_start And Cells(i, 1).Value <> Cells(i - 1, 1).Value Then
    Start_price = Cells(i, 3).Value
    
    ElseIf Cells(i, 1).Value <> Cells(i + 1, 1).Value And Cells(i, 2).Value = year_end Then
    End_price = Cells(i, 3).Value
    
    Range("J" & ticker_row_price).Value = End_price - Start_price
    Range("K" & ticker_row_price).Value = ((End_price - Start_price) / Start_price)
    ticker_row_price = ticker_row_price + 1
    
    ElseIf Cells(i, 1).Value <> Cells(i + 1, 1).Value And Cells(i, 2).Value <> year_end Then
    End_price = Cells(i, 3).Value
    Range("J" & ticker_row_price).Value = End_price - Start_price
    Range("K" & ticker_row_price).Value = ((End_price - Start_price) / Start_price)
    ticker_row_price = ticker_row_price + 1
    
    End If

Next i

'Formatting yearly change percentages
Range("K2").Select
Range(Selection, Selection.End(xlDown)).Select
Selection.NumberFormat = "0.00%"

Range("J1").Select
year_rows = Range("J1", Selection.End(xlDown)).Rows.Count

For i = 2 To year_rows
    If Cells(i, 11).Value < 0 Then
    Cells(i, 11).Interior.ColorIndex = 3
    Else
    Cells(i, 11).Interior.ColorIndex = 4
    End If

Next i
    
End Sub