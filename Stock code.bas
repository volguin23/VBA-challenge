Attribute VB_Name = "Module1"
Sub stock_data()

 For Each ws In Worksheets

ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "Total Stock Volume"
ws.Cells(1, 16).Value = "Ticker"
ws.Cells(1, 17).Value = "Value"
ws.Cells(2, 15).Value = "Greatest % Increase"
ws.Cells(3, 15).Value = "Greatest % Decrease"
ws.Cells(4, 15).Value = "Greatest Total Volume"


ws.Range("k2:k753003").NumberFormat = "0.00%"
ws.Cells(2, 17).NumberFormat = "0.00%"
ws.Cells(3, 17).NumberFormat = "0.00%"


  ' Set an initial variable for holding the ticker name
  Dim ticker_Name As String

  ' Set an initial variable for holding the yearly change
  Dim yearly_change As Double
  Dim opening As Double
  Dim closing As Double
  yearly_change = 0
  opening = ws.Cells(2, 3).Value
  closing = 0
  
  
  ' Set an initial variable for holding the percent change
  Dim percent_change As Double
    Dim opening_change As Double
  Dim closing_change As Double
percent_change = 0
  opening_change = ws.Cells(2, 3).Value
  closing_change = 0
  
  
  ' Set an initial variable for holding the total stock volume
  Dim total_vlm As Double
  total_vlm = 0

  ' Keep track of the location for each ticker in the summary table
  Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2

  ' Loop through all tickers
  For i = 2 To 753003

    ' Check if we are still within the same ticker, if it is not...
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

      ' Set the ticker
      ticker_Name = ws.Cells(i, 1).Value
      
   ' Set the closing value
     closing = ws.Cells(i, 6).Value
'MsgBox (closing)

'set the closing change value
 closing_change = ws.Cells(i, 6).Value
      ' Add to the Total Volume
      total_vlm = total_vlm + ws.Cells(i, 7).Value
      
           ' get the yearly change
    yearly_change = closing - opening
    'MsgBox (yearly_change)
      
      'Get the percent change
      percent_change = (1 - closing_change / opening_change) * -1
      
      ' Print the ticker in the Summary Table
      ws.Range("i" & Summary_Table_Row).Value = ticker_Name

      ' Print the total volume to the Summary Table
      ws.Range("L" & Summary_Table_Row).Value = total_vlm
         ' Print the yeary change to the Summary Table
      ws.Range("j" & Summary_Table_Row).Value = yearly_change
      'print the percent change to the Summary Table
      ws.Range("k" & Summary_Table_Row).Value = percent_change
      
      ' Add one to the summary table row
      Summary_Table_Row = Summary_Table_Row + 1
      
      ' Reset the Total Volume
      total_vlm = 0
  
      'Set the opening value again
opening = ws.Cells(i + 1, 3).Value
 'MsgBox (opening)

'set the opening changes value again
opening_change = ws.Cells(i + 1, 3).Value

      
    ' If the cell immediately following a row is the same ticker...
    Else
    
      ' Add to the Total Volume
   total_vlm = total_vlm + ws.Cells(i, 7).Value
   

   'Reset the closing
      'closing = 0

  ' Reset the opening
     'opening = 0


    End If
    
   
  Next i

Dim function_a As Double
Dim function_b As Double
Dim function_c As String
Dim function_d As String
Dim function_e As String

'call function to increase
function_a = getmaxvalue("11")
' print value
ws.Cells(2, 17).Value = function_a


'call function to increase string
function_c = getmaxvalueticker("11")
' print value
ws.Cells(2, 16).Value = function_c

'call function to decrease
function_b = getminvalue("11")
' print value
ws.Cells(3, 17).Value = function_b

'call function to decrease string
function_e = getminvalueticker("11")
' print value
ws.Cells(3, 16).Value = function_e

'call function to total volume
function_a = getmaxvalue("12")
' print value
ws.Cells(4, 17).Value = function_a

'call function to total string
function_d = getmaxvaluetickerTotal("12")
' print value
ws.Cells(4, 16).Value = function_d

    Next ws

End Sub

Function getmaxvalue(range_a As Double)
Dim higher
 
'For Each ws In Worksheets

 'getmaxvalue = 0
 
For k = 2 To 3001

If Cells(k, range_a).Value > higher Then
higher = Cells(k, range_a).Value
End If
Next k
getmaxvalue = higher
 'Cells(2, 17).Value = getmaxvalue
 
'Next ws



End Function

Function getminvalue(range_b As Double)
'For Each ws In Worksheets

Dim lower

For l = 2 To 3001

If Cells(l, range_b).Value < lower Then
lower = Cells(l, range_b).Value

End If
Next l

getminvalue = lower

'Next ws
End Function

Function getmaxvalueticker(range_c As Double)
'For Each ws In Worksheets

Dim higherString As String
Dim higherdos As Double
For m = 2 To 3001

If Cells(m, range_c).Value > higherdos Then
higherdos = Cells(m, range_c).Value
higherString = Cells(m, range_c - 2).Value
End If
Next m

getmaxvalueticker = higherString

'Next ws

End Function

Function getmaxvaluetickerTotal(range_d As Double)
'For Each ws In Worksheets

Dim higherStringTotal As String
Dim higherdosTotal As Double
For n = 2 To 3001

If Cells(n, range_d).Value > higherdosTotal Then
higherdosTotal = Cells(n, range_d).Value
higherStringTotal = Cells(n, range_d - 3).Value
End If
Next n

getmaxvaluetickerTotal = higherStringTotal

'Next ws

End Function

Function getminvalueticker(range_e As Double)
'For Each ws In Worksheets

Dim lowerString As String
Dim lowerdos As Double
For o = 2 To 3001

If Cells(o, range_e).Value < lowerdos Then
lowerdos = Cells(o, range_e).Value
lowerString = Cells(o, range_e - 2).Value
End If
Next o

getminvalueticker = lowerString

'Next ws

End Function

