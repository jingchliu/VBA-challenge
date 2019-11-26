Sub stockfilter():

Dim lastRow As Long
Dim lastRowTicker As Long
Dim ticker As String
Dim yearlyChange As Double
Dim percentageChange As Double
Dim totalVolumn As Double
Dim summaryRow As Integer
Dim closePrice As Double
Dim openPrice As Double
Dim lastRowResult As Long
Dim max As Double
Dim min As Double
Dim maxTotal As Double
summaryRow = 2
lastRowTicker = 1
totalVolumn = 0

For Each ws In Worksheets
    summaryRow = 2
    lastRowTicker = 1
    totalVolumn = 0
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percentage Change"
    ws.Cells(1, 12).Value = "Total Stock Volumn"
    lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volumn"
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"

    For i = 2 To lastRow
        totalVolumn = totalVolumn + ws.Cells(i, 7).Value
        If ws.Cells(i + 1, 1).Value = ws.Cells(i, 1).Value Then
            lastRowTicker = lastRowTicker + 1
        End If
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            ticker = ws.Cells(i, 1).Value
            openPrice = ws.Cells(i - lastRowTicker + 1, 3).Value
            closePrice = ws.Cells(i, 6).Value
            yearlyChange = closePrice - openPrice
            If openPrice<>0  Then
                percentageChange = yearlyChange / openPrice
            End If
            ws.Cells(summaryRow, 9).Value = ticker
            ws.Cells(summaryRow, 10).Value = yearlyChange
            ws.Cells(summaryRow, 11).Value = percentageChange
            ws.Cells(summaryRow, 11).NumberFormat = "0.00%"
            ws.Cells(summaryRow, 12).Value = totalVolumn
                If percentageChange > 0 Then
                    ws.Cells(summaryRow, 11).Interior.ColorIndex = 4
                ElseIf percentageChange < 0 Then
                    ws.Cells(summaryRow, 11).Interior.ColorIndex = 3
                End If
            summaryRow = summaryRow + 1
            lastRowTicker = 1
            totalVolumn = 0
       End If
   Next i
   
    lastRowResult = ws.Cells(Rows.Count, 9).End(xlUp).Row
    min = ws.Cells(2, 11).Value
    max = ws.Cells(2, 11).Value
    maxTotal = ws.Cells(2, 12).Value
        For j = 3 To lastRowResult
            If ws.Cells(j, 11).Value > max Then
                max = ws.Cells(j, 11).Value
                ws.Cells(2, 17).Value = max
                ws.Cells(2, 16).Value = ws.Cells(j, 9).Value
            End If
        
            If ws.Cells(j, 11).Value < min Then
                min = ws.Cells(j, 11).Value
                ws.Cells(3, 17).Value = min
                ws.Cells(3, 16).Value = ws.Cells(j, 9).Value
            End If
        
            If ws.Cells(j, 12).Value > maxTotal Then
                maxTotal = ws.Cells(j, 12).Value
                ws.Cells(4, 17).Value = maxTotal
                ws.Cells(4, 16).Value = ws.Cells(j, 9).Value
            End If
        ws.Range("Q2:Q3").NumberFormat = "0.00%"
        Next j

Next ws

End Sub

