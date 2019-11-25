Sub stockfilter():

Dim lastRow As Long
Dim lastRowTicker As Long
Dim ticker As String
Dim yearlyChange As Double
Dim percentageChange As Double
Dim totalVolumn As Long
Dim summaryRow As Integer

totalVolumn=0
summaryRow=2
lastRowTicker=1

For Each ws In Worksheets
    ws.Cells(1,9).Value="Ticker"
    ws.Cells(1,10).Value="Yearly Change"
    ws.Cells(1,11).Value="Percentage Change"
    ws.Cells(1,12).Value="Total Stock Volumn"
    lastRow=ws.Cells(Rows.Count,1).End(xlUp).Row

    For i=2 to lastRow
        If ws.Cells(i,1).Value=ws.Cells(i+1,1).Value Then
            lastRowTicker=lastRowTicker+1
        End If

        If ws.Cells(i,1).Value<>ws.Cells(i-1,1).Value Then
            
                ticker=ws.Cells(i,1).Value
                yearlyChange=ws.Cells(i-1,6).Value-ws.Cells(i-lastRowTicker,3).Value
                percentageChange=yearlyChange/ws.Cells(i-lastRowTicker,3)
                totalVolumn=totalVolumn+ws.Cells(i-1,7).Value
            
                ws.Cells(summaryRow,9).Value=ticker
                ws.Cells(summaryRow,10).Value=yearlyChange
                ws.Cells(summaryRow,11).Value=percentageChange
                ws.Cells(summaryRow,11).NumberFormat="0.00%"
                ws.Cells(summaryRow,12).Value=totalVolumn
                    If percentageChange>0 Then
                        ws.Cells(summaryRow,11).Interior.ColorIndex=4
                    Elseif percentageChange<0 Then
                        ws.Cells(summaryRow,11).Interior.ColorIndex=3
                    End If
                summaryRow=summaryRow+1
                totalVolumn=0
                lastRowTicker=1
            

        End If

    Next i

Next ws

End Sub