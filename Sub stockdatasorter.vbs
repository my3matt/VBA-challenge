Sub stockdatasorter()
'Define Dimensions
Dim ws As Worksheet
Dim ticker As String
Dim vol As Integer
Dim year_Open As Double
Dim year_close As Double
Dim yearly_change As Double
Dim percent_change As Double
Dim Summary_Table_Row As Integer

'Overflow Error and compute results to all worksheets
On Error Resume Next
For Each ws In ThisWorkbook.Worksheets

'Define Header Columns
ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "Total Stock Volume"

'Using Summary Table to compute data simultaneously
Summary_Table_Row = 2

'Loop data through worksheet
For i = 2 To ws.UsedRange.Rows.Count

'Calculate values to compute in respective headers
If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
ticker = ws.Cells(i, 1).Value
vol = ws.Cells(i, 7).Value
year_Open = ws.Cells(i, 3).Value
year_close = ws.Cells(i, 6).Value
yearly_change = year_close - year_Open
percent_change = year_close / year_Open

'Fill header cells with calculated values
ws.Cells(Summary_Table_Row, 9).Value = ticker
ws.Cells(Summary_Table_Row, 10).Value = yearly_change
ws.Cells(Summary_Table_Row, 11).Value = percent_change
ws.Cells(Summary_Table_Row, 12).Value = vol
Summary_Table_Row = Summary_Table_Row + 1
vol = 0

'Color Conditionals
With ws.Range("J", 10).Value.FormatConditions.Add(xlCellValue, xlLess, "=0")
                            .Interior.ColorIndex = 3
                            End With
With ws.Range("J", 10).Value.FormatConditions.Add(xlCellValue, xlGreater, "=0")
                            .Interior.ColorIndex = 4
                            End With

End If
Next i
'Entering percent change in correct format
ws.Columns("K").NumberFormat = "0.00%"
Next
End Sub


