Sub stockMarket()
' Homework2 - The VBA of Wall Street
' Made by Imara Paiz on 10/29/2017
' Data Analytic Bootcamp 
' UC Berkeley Extension
Dim i, j As Double
Dim LastRow As Double
Dim totVol As Double
Dim valueOpen, valueClose, yearlyChange As Double
Dim percentChange, MaxIncrease, MaxDecrease, MaxTotVol As Double
Dim currentYear, dateOpen, dateClose As String
Dim n As Integer
n = 1

' Loop trough all sheets
For Each ws In Worksheets

    ' calculate last row number in the current worksheet
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    totVol = 0
    j = 2

    ' set all header titles
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Value"
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"
    
    ' calculate total stock value in the current worksheet
    For i = 2 To LastRow
        totVol = totVol + ws.Cells(i, 7).Value
        If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
           ws.Cells(j, 9).Value = ws.Cells(i, 1).Value
           ws.Cells(j, 12).Value = totVol
           totVol = 0
           j = j + 1
        End If
    Next i
    
    'calculate yearly change in the current worksheet
    j = 2
    valueOpen = 0#
    valueClose = 0#
    yearlyChange = 0#
    percentChange = 0#
    valueOpen = CDbl(ws.Cells(2, 3).Value)
    For i = 2 To LastRow
        If (ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value) Then
            valueClose = CDbl(ws.Cells(i, 6).Value)
            yearlyChange = valueClose - valueOpen
            If yearlyChange < 0# Then
                ws.Cells(j, 10).Interior.ColorIndex = 3
            Else
                ws.Cells(j, 10).Interior.ColorIndex = 4
            End If
            ws.Cells(j, 10).Value = yearlyChange
            
            ' Calculate the percentChange value
            If valueOpen > 0# Then
                percentChange = yearlyChange / valueOpen
            End If
            ws.Cells(j, 11).Value = percentChange
            ws.Cells(j, 11).NumberFormat = "####0.00%"
            valueOpen = CDbl(ws.Cells(i + 1, 3).Value)
            j = j + 1
        End If
    Next i
    n = n + 1
    
    'Calculate Greatest percentage Increase, Greatest percentage Decrease and Greatest Total Volume
    MaxIncrease = 0
    MaxDecrease = 0
    MaxTotVol = 0
    
    For i = 2 To j
        If CDbl(ws.Cells(i, 11).Value) > MaxIncrease Then
            MaxIncrease = CDbl(ws.Cells(i, 11).Value)
            MaxTicker = Cells(i, 9).Value
        End If
        If CDbl(ws.Cells(i, 11).Value) < MaxDecrease Then
            MaxDecrease = CDbl(ws.Cells(i, 11).Value)
            MinTicker = ws.Cells(i, 9).Value
        End If
        If CDbl(Cells(i, 12).Value) > MaxTotVol Then
            MaxTotVol = CDbl(ws.Cells(i, 12).Value)
            MaxTotTicker = ws.Cells(i, 9).Value
        End If
    Next i
    ws.Cells(2, 16).Value = MaxTicker
    ws.Cells(2, 17).Value = MaxIncrease
    ws.Cells(2, 17).NumberFormat = "####0.00%"
    ws.Cells(3, 16).Value = MinTicker
    ws.Cells(3, 17).Value = MaxDecrease
    ws.Cells(3, 17).NumberFormat = "####0.00%"
    ws.Cells(4, 16).Value = MaxTotTicker
    ws.Cells(4, 17).Value = MaxTotVol
Next ws

End Sub

Sub clearCalculations()
Dim LastRow As Double

For Each ws In Worksheets
  LastRow = ws.Cells(Rows.Count, 9).End(xlUp).Row
  For i = 1 To LastRow
    ws.Cells(i, 9).Value = ""
    ws.Cells(i, 10).Interior.ColorIndex = 0
    ws.Cells(i, 10).Value = ""
    ws.Cells(i, 11).Value = ""
    ws.Cells(i, 12).Value = ""
    ws.Cells(i, 13).Value = ""
    ws.Cells(i, 14).Value = ""
    ws.Cells(i, 15).Value = ""
    ws.Cells(i, 16).Value = ""
    ws.Cells(i, 17).Value = ""
  Next i
Next ws
End Sub


