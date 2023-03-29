Attribute VB_Name = "Module1"
Sub stock_market()

For Each ws In Worksheets

Dim WorksheetName As String
WorksheetName = ws.Name

ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "Total Stock Volume"
ws.Cells(2, 15).Value = "Greatest % Increase"
ws.Cells(3, 15).Value = "Greatest % Decrease"
ws.Cells(4, 15).Value = "Greatest Total Volume"
ws.Cells(1, 16).Value = "Ticker"
ws.Cells(1, 17).Value = "Value"

Dim Ticker As String
Dim Yearly_Change As Double
Dim Percent_Change As Double
Dim Total_Stock As Long
Dim Initial_Table_Row As Integer
Dim Summary_Table_Row As Integer
Dim LastRow1 As Long
Dim i As Long
Dim j As Long

Initial_Table_Row = 2
Summary_Table_Row = 2
j = 2
LastRow1 = ws.Cells(Rows.Count, 1).End(xlUp).Row

For i = 2 To LastRow1

If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
ws.Cells(Initial_Table_Row, 9).Value = ws.Cells(i, 1).Value

Yearly_Change = ws.Cells(i, 6).Value - ws.Cells(j, 3).Value
ws.Range("J" & Summary_Table_Row).Value = Yearly_Change
    If Yearly_Change < 0 Then
    ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
    Else
    ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
    End If
    If ws.Cells(j, 3).Value <> 0 Then
    Percent_Change = (ws.Cells(i, 6).Value - ws.Cells(j, 3).Value) / ws.Cells(j, 3).Value
    ws.Range("K" & Summary_Table_Row).Value = Percent_Change
    ws.Range("K" & Summary_Table_Row).Value = Format(Percent_Change, "Percent")
    End If
    

ws.Cells(Initial_Table_Row, 12).Value = WorksheetFunction.Sum(Range(ws.Cells(j, 7), ws.Cells(i, 7)))
ws.Range("L" & Summary_Table_Row).Value = Total_Stock

Initial_Table_Row = Initial_Table_Row + 1
Summary_Table_Row = Summary_Table_Row + 1
j = i + 1

End If

Next i

Dim LastRow2 As Long
Dim Greatest_Increase As Double
Dim Greatest_Decrease As Double
Dim Greatest_Volume As Double

LastRow2 = ws.Cells(Rows.Count, 9).End(xlUp).Row
Greatest_Increase = ws.Cells(2, 11).Value
Greatest_Decrease = ws.Cells(2, 11).Value
Greatest_Volume = ws.Cells(2, 12).Value

For i = 2 To LastRow2
    If ws.Cells(i, 11).Value > Greatest_Increase Then
    Greatest_Increase = ws.Cells(i, 11).Value
    ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
    ws.Cells(2, 17).Value = Format(Greatest_Increase, "Percent")
    End If
    If ws.Cells(i, 11).Value < Greatest_Decrease Then
    Greatest_Decrease = ws.Cells(i, 11).Value
    ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
    ws.Cells(3, 17).Value = Format(Greatest_Decrease, "Percent")
    End If
    If ws.Cells(i, 12).Value > Greatest_Volume Then
    Greatest_Volume = ws.Cells(i, 12).Value
    ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
    ws.Cells(4, 17).Value = Format(Greatest_Volume, "Scientific")
    End If
    
Next i
Worksheets(WorksheetName).Columns("A:Z").AutoFit
Next ws

End Sub
