Attribute VB_Name = "Module1"
Sub VBA_Assignment():

For Each ws In Worksheets
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"

Dim Ticker As String
Dim openYearly As Double
Dim totalVolume As Double
totalVolume = 0
Dim totalYearly As Double
totalYearly = 0
Dim percentChange As Double
Dim tickerRow As Long
tickerRow = 2
Dim lastRow As Long
lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

For i = 2 To lastRow
openYearly = ws.Cells(tickerRow, 3).Value

    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        Ticker = ws.Cells(i, 1).Value
        ws.Range("I" & tickerRow).Value = Ticker
    
        totalYearly = totalYearly + (ws.Cells(i, 6).Value - openYearly)
        ws.Range("J" & tickerRow).Value = totalYearly
    
        percentChange = (totalYearly / openYearly)
        ws.Range("K" & tickerRow).Value = percentChange
        ws.Range("K" & tickerRow).Style = "Percent"
        
        totalVolume = totalVolume + ws.Cells(i, 7).Value
        ws.Range("L" & tickerRow).Value = totalVolume
        
       
        tickerRow = tickerRow + 1
        totalYearly = 0
        totalVolume = 0
        openYearly = ws.Cells(tickerRow, 3).Value
    
    Else
        totalVolume = totalVolume + ws.Cells(i, 7).Value
    
    End If
Next i

For i = 2 To lastRow
    If ws.Cells(i, 10).Value >= 0 Then
        ws.Cells(i, 10).Interior.ColorIndex = 4
    Else
        ws.Cells(i, 10).Interior.ColorIndex = 3
    End If

Next i
    
Next ws

End Sub

