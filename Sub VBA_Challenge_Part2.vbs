Sub VBA_Challenge_Part2():

'Set variables
Dim Ticker As String
Dim Yearly_Change As Double
Dim Percentage_Change As Double
Dim Total_Stock_Volume As Double
Dim Summary_Table_Row As Long
Dim Trading_Days As Long
Dim LastRow As Long

'Loop through all worksheets
Dim ws As Worksheet
For Each ws In Worksheets

'Set variable opening values
Total_Stock_Volume = 0
Summary_Table_Row = 2
Trading_Days = 0
LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

'Insert Summary Table Headers
ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Percentage Change"
ws.Cells(1, 12).Value = "Total Stock Volume"

'clear formatting ticker row
ws.Range("a1:g" & LastRow).ClearFormats

For I = 2 To LastRow

'Find where new ticker symbol starts on current worksheet (f we are not in the same ticker, then..)
If ws.Cells(I + 1, 1).Value <> ws.Cells(I, 1).Value Then

'Message Box the unique Ticker Symbol as check
'MsgBox (Cells(i, 1).Value)
'Set the ticker Symbol
Ticker = ws.Cells(I, 1).Value

'Count trading day
Trading_Days = Trading_Days + 1

' Calculate Yearly Change, % Change and Total Stock Volume
Yearly_Change = ws.Cells(I, 6).Value - ws.Cells(I - (Trading_Days - 1), 3).Value
Total_Stock_Volume = Total_Stock_Volume + ws.Cells(I, 7).Value
If ws.Cells(I - (Trading_Days - 1), 3).Value <> 0 Then
Percentage_Change = (ws.Cells(I, 6).Value - ws.Cells(I - (Trading_Days - 1), 3).Value) / ws.Cells(I - (Trading_Days - 1), 3).Value
Else
Percentage_Change = 0
End If
'Print to the Summary Table on each page
ws.Range("I" & Summary_Table_Row).Value = Ticker
ws.Range("J" & Summary_Table_Row).Value = Yearly_Change
ws.Range("K" & Summary_Table_Row).Value = Percentage_Change
ws.Range("L" & Summary_Table_Row).Value = Total_Stock_Volume
      
      
      
'Set up next loop - add one to the summary table row and reset variables
Summary_Table_Row = Summary_Table_Row + 1
Yearly_Change = 0
Total_Stock_Volume = 0
Trading_Days = 0
      
Else
'Add to Total Stock Volume
Total_Stock_Volume = Total_Stock_Volume + ws.Cells(I, 7).Value
      
 'Count trading day
 Trading_Days = Trading_Days + 1

End If

Next I




'Apply conditional formatting to summary table
Dim LastRowSummary As Integer
LastRowSummary = ws.Cells(Rows.Count, 10).End(xlUp).Row

For j = 2 To LastRowSummary
    
'% for percentage change
ws.Cells(j, 11).NumberFormat = "0.00%"
ws.Cells(j, 10).NumberFormat = "0.00"

'autofit columns
ws.Range("I:L").EntireColumn.AutoFit

'red background for less than zero and green background for greater than zero
If ws.Cells(j, 10).Value < 0 Then
ws.Cells(j, 10).Interior.ColorIndex = 3
    
ElseIf ws.Cells(j, 10).Value > 0 Then
ws.Cells(j, 10).Interior.ColorIndex = 4

Else

End If
Next j

'Set Variables for Second Summary Table
Dim Greatest_Increase As Double
Dim Greatest_Decrease As Double
Dim Greatest_Volume As Double
Dim RangePercent As Range
Dim RangeVolume As Range
Dim LastRowSum As Integer


'Find last row
LastRowSum = ws.Cells(Rows.Count, 11).End(xlUp).Row

'Print labels for second summary table
ws.Cells(2, 15).Value = "Greatest % Increase"
ws.Cells(3, 15).Value = "Greatest % Decrease"
ws.Cells(4, 15).Value = "Greatest Stock Volume"
ws.Cells(1, 16).Value = "Ticker"
ws.Cells(1, 17).Value = "Value"

'Set range from which to determine values
Set RangePercent = ws.Range("K2:K" & LastRowSum)
Set RangeVolume = ws.Range("L2:L" & LastRowSum)

'Find values for greatest % increase, greatest % decrease and greatest volume
Greatest_Increase = ws.Application.WorksheetFunction.Max(RangePercent)
Greatest_Decrease = ws.Application.WorksheetFunction.Min(RangePercent)
Greatest_Volume = ws.Application.WorksheetFunction.Max(RangeVolume)

'Print to table
ws.Cells(2, 17).Value = Greatest_Increase
ws.Cells(3, 17).Value = Greatest_Decrease
ws.Cells(4, 17).Value = Greatest_Volume

'% for percentage change
ws.Range("Q2:Q3").NumberFormat = "0.00%"

'Print ticker to table
For I = 2 To LastRowSum

If ws.Cells(I, 11).Value = Greatest_Increase Then
ws.Cells(2, 16).Value = ws.Cells(I, 9).Value

ElseIf ws.Cells(I, 11).Value = Greatest_Decrease Then
ws.Cells(3, 16).Value = ws.Cells(I, 9).Value

ElseIf ws.Cells(I, 12).Value = Greatest_Volume Then
ws.Cells(4, 16).Value = ws.Cells(I, 9).Value

Else

End If

Next I

'autofit columns
ws.Range("O:Q").EntireColumn.AutoFit

Next ws

End Sub


