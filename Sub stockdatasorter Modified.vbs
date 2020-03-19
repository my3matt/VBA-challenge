Sub stockdatasorter()
    'Loop through each worksheet
Dim ws As Worksheet
    For Each ws In ActiveWorkbook.Worksheets
    ws.Activate
        'Determine Last row
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    'Define Dimensions
        Dim year_Open As Double
        Dim year_close As Double
        Dim yearly_change As Double
        Dim ticker As String
        Dim Percent_Change As Double
        Dim Volume As Double
        Volume = 0
        Dim Row As Double
        Row = 2
        Dim Column As Integer
        Column = 1
        Dim i As Long

        ' Define Header Columns
        Cells(1, "I").Value = "Ticker"
        Cells(1, "J").Value = "Yearly Change"
        Cells(1, "K").Value = "Percent Change"
        Cells(1, "L").Value = "Total Stock Volume"

        
        'Set Year Open
        year_Open = Cells(2, Column + 2).Value
         
         ' Loop data through worksheet
        For i = 2 To LastRow
         
         ' If function to compute in respective headers
            If Cells(i + 1, Column).Value <> Cells(i, Column).Value Then
                ' Set Ticker
                ticker = Cells(i, Column).Value
                Cells(Row, Column + 8).Value = ticker
                ' Set Year Close
                year_close = Cells(i, Column + 5).Value
                ' Calculate for Yearly Change
                yearly_change = year_close - year_Open
                Cells(Row, Column + 9).Value = yearly_change
                ' Calculate for percent change
                '-------if function to account for div/0
                If (year_Open = 0 And year_close = 0) Then
                    Percent_Change = 0
                ElseIf (year_Open = 0 And year_close <> 0) Then
                    Percent_Change = 1
                'Calculate for percent change and format it accordingly
                Else
                    Percent_Change = yearly_change / year_Open
                    Cells(Row, Column + 10).Value = Percent_Change
                    Cells(Row, Column + 10).NumberFormat = "0.00%"
                End If
                ' Add Total Volume
                Volume = Volume + Cells(i, Column + 6).Value
                Cells(Row, Column + 11).Value = Volume
                ' Add one to the summary table row
                Row = Row + 1
                ' reset open
                year_Open = Cells(i + 1, Column + 2)
                ' reset the Volume Total
                Volume = 0
            'if cells are the same ticker
            Else
                Volume = Volume + Cells(i, Column + 6).Value
            End If
        Next i
        
        ' Determine the Last Row of Yearly Change per WS
        YCLastRow = ws.Cells(Rows.Count, Column + 8).End(xlUp).Row
        ' Color conditionals
        For j = 2 To YCLastRow
            If (Cells(j, Column + 9).Value > 0 Or Cells(j, Column + 9).Value = 0) Then
                Cells(j, Column + 9).Interior.ColorIndex = 4
            ElseIf Cells(j, Column + 9).Value < 0 Then
                Cells(j, Column + 9).Interior.ColorIndex = 3
            End If
        Next j
        
   
        
    Next ws
        
End Sub

