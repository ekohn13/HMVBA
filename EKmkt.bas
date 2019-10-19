Attribute VB_Name = "EKsmt"
Sub smt()
For Each ws In Sheets 'Object ws allows us to work on all worksheets
Dim counting As Integer   'This variable is used to move and add value to a certain column
Dim Stock_open As Double
Dim Stock_close As Double
Dim Yearly_Change As Double
Dim Percent_Change As Double
Dim Stock_Volume As Single
Dim Decrease_Sticker As String
Dim Volume_Sticker As String
'This Variables are used to create a funtion that will find a Max value in a Range
Dim rng As Range
Dim dblMax As Double
'Headers
Range("I1").Value = "Ticker"
Range("J1").Value = "Year Change"
Range("K1").Value = "Percent Change"
Range("L1").Value = "Volume"
Range("P1").Value = "Ticker"
Range("Q1").Value = "Value"
Range("O2").Value = "Greatest % Increase"
Range("O3").Value = "Greatest % Decrease"
Range("O4").Value = "Greatest Total Volume"
'Initializing Variables
counting = 2
Stock_Volume = 0
Greatest_Increase = 0
Stock_open = Range("C2").Value
'Finding the last row to set the limit in the for loop
LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
'This loop will iterate through the data to find the percent change, the year change and volume of each stock
For i = 2 To LastRow
  If Cells(i + 1, 1).Value = Cells(i, 1).Value Then  'Finding the total volumen in each stock
  Stock_Volume = Stock_Volume + Cells(i, 7).Value
  'Finding Year Change in each stock
  ElseIf Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    Cells(counting, 9).Value = Cells(i, 1)
    Stock_close = Cells(i, 6).Value
    Yearly_Change = Stock_close - Stock_open
    Cells(counting, 10).Value = Yearly_Change
    If Yearly_Change > 0 Then                      'Formatting the cell color according to its value
    Cells(counting, 10).Interior.ColorIndex = 4
    Else:
    Cells(counting, 10).Interior.ColorIndex = 3
    End If
    If Stock_open > 0 Then
    Percent_Change = (Stock_close - Stock_open) / Stock_open
    Else
    Percent_Change = Stock_open
    End If
    
    Cells(counting, 11).Value = FormatPercent(Percent_Change, 2)
    
    Cells(counting, 12).Value = Stock_Volume
    Stock_Volume = 0
    
    Stock_open = Cells(i + 1, 3).Value
    counting = counting + 1
  End If
Next i
Range("L291").Value = " "
'Finding the Greatest % increase, decrease and greatest volume
LastRow = Cells(Rows.Count, "K").End(xlUp).Row
For j = 2 To LastRow
  If Cells(j + 1, 11).Value > Cells(j, 11).Value And Greatest_Increase < Cells(j + 1, 11).Value Then
   Greatest_Increase = Cells(j + 1, 11).Value
   Greatest_Sticker = Cells(j + 1, 9).Value
  End If
Next j
Range("Q2").Value = FormatPercent(Greatest_Increase, 2)
Range("P2").Value = Greatest_Sticker
For k = 2 To LastRow
  If Cells(k + 1, 11).Value < Cells(k, 11).Value And Greatest_Decrease > Cells(k + 1, 11).Value Then
   Greatest_Decrease = Cells(k + 1, 11).Value
   Decrease_Sticker = Cells(k + 1, 9).Value
  End If
Next k
Range("Q3").Value = FormatPercent(Greatest_Decrease, 2)
Range("P3").Value = Decrease_Sticker
Set rng = ws.Range("L2:l" & LastRow)
dblMax = Application.WorksheetFunction.Max(rng)
ws.Range("Q4").Value = dblMax
For i = 2 To LastRow
  If Cells(i, 12) = dblMax Then
  ws.Range("P4").Value = Cells(i, 9)
  End If
Next i

Next ws
End Sub


