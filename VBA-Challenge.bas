Attribute VB_Name = "Module1"
Sub stockmarketchallenge()

For Each Sheet In ThisWorkbook.Worksheets

Sheet.Cells(1, 9).Value = "Ticker"
Sheet.Cells(1, 10).Value = "Yearly Change"
Sheet.Cells(1, 11).Value = "Percent Change"
Sheet.Cells(1, 12).Value = "Total Stock Volume"
Sheet.Cells(2, 15).Value = "Greatest%Increase"
Sheet.Cells(3, 15).Value = "Greatest%Decrease"
Sheet.Cells(4, 15).Value = "Greatest Total Volume"
Sheet.Cells(1, 16).Value = "Ticker"
Sheet.Cells(1, 17).Value = "Value"

Dim i As Long
Dim j As Long
Dim lastrow1 As Long
Dim tickernumber As Long
Dim PercentChange As Double

lastrow1 = Sheet.Cells(Rows.Count, 1).End(xlUp).Row
j = 2
tickernumber = 2

For i = 2 To lastrow1

If Sheet.Cells(i, 1).Value <> Sheet.Cells(i + 1, 1).Value Then
Sheet.Cells(tickernumber, 9).Value = Sheet.Cells(i, 1).Value
Sheet.Cells(tickernumber, 10).Value = Sheet.Cells(i, 6).Value - Sheet.Cells(j, 3).Value
    If Sheet.Cells(tickernumber, 10).Value > 0 Then
    Sheet.Cells(tickernumber, 10).Interior.ColorIndex = 4
    Else
    Sheet.Cells(tickernumber, 10).Interior.ColorIndex = 3
    End If
    
    If Sheet.Cells(j, 3).Value <> 0 Then
    PercentChange = ((Sheet.Cells(i, 6).Value - Sheet.Cells(j, 3).Value) / Sheet.Cells(j, 3).Value)
    Sheet.Cells(tickernumber, 11).Value = Format(PercentChange, "Percent")
    End If
Sheet.Cells(tickernumber, 12).Value = WorksheetFunction.Sum(Range(Sheet.Cells(j, 7), Sheet.Cells(i, 7)))
j = i + 1
tickernumber = tickernumber + 1
    End If
    
Next i

Dim lastrow2 As Long
Dim GreatestIncrease As Double
Dim GreatestDecrease As Double
Dim HighestStock As Double

lastrow2 = Sheet.Cells(Rows.Count, 9).End(xlUp).Row
GreatestIncrease = Sheet.Cells(2, 11).Value
GreatestDecrease = Sheet.Cells(2, 11).Value
HighestStock = Sheet.Cells(2, 12).Value

For i = 2 To lastrow2

    If Sheet.Cells(i, 11).Value > GreatestIncrease Then
    GreatestIncrease = Sheet.Cells(i, 11).Value
    Sheet.Cells(2, 16).Value = Sheet.Cells(i, 9).Value
    Sheet.Cells(2, 17).Value = Format(GreatestIncrease, "Percent")
    End If
    
    If Sheet.Cells(i, 11).Value < GreatestDecrease Then
    GreatestDecrease = Sheet.Cells(i, 11).Value
    Sheet.Cells(3, 16).Value = Sheet.Cells(i, 9).Value
    Sheet.Cells(3, 17).Value = Format(GreatestDecrease, "Percent")
    End If
    
    If Sheet.Cells(i, 12).Value > HighestStock Then
    HighestStock = Sheet.Cells(i, 12).Value
    Sheet.Cells(4, 16).Value = Sheet.Cells(i, 9).Value
    Sheet.Cells(4, 17).Value = Format(HighestStock, "Scientific")
    End If
        
    Next i
          
Next Sheet

End Sub

