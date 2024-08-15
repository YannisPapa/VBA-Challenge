Attribute VB_Name = "Module2"
Sub ticker()

    Dim ws As Worksheet

    'loop over each sheet
    For Each ws In Worksheets
        'Create Variables
        Dim i As Long
        Dim j As Integer
        Dim numRows As Long
        Dim countTicker As Integer
        Dim ticker As String
        Dim openPrice As Double
        Dim closePrice As Double
        Dim totalVolume As Double
        
        'Initialize Variables
        numRows = ws.Cells(Rows.Count, 1).End(xlUp).Row
        countTicker = 2
        ticker = ws.Range("A2").Value
        openPrice = ws.Range("C2").Value
        closePrice = 0
        totalVolume = 0
        
        'Print Headers
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Quarterly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        
        ws.Range("O1").Value = "Ticker"
        ws.Range("P1").Value = "Value"
        ws.Range("N2").Value = "Greatest % Increase"
        ws.Range("N3").Value = "Greatest % Decrease"
        ws.Range("N4").Value = "Greatest Total Volume"
        
        For i = 2 To numRows
            If ticker = ws.Cells(i, 1).Value Then
                'calculate volume and get final closing price
                totalVolume = totalVolume + ws.Cells(i, 7).Value
                closePrice = ws.Cells(i, 6).Value
            Else
                'print data
                ws.Cells(countTicker, 9).Value = ticker
                ws.Cells(countTicker, 10).Value = closePrice - openPrice
                ws.Cells(countTicker, 11).Value = ((closePrice - openPrice) / openPrice)
                ws.Cells(countTicker, 12).Value = totalVolume
                'Reset Variables
                countTicker = countTicker + 1
                ticker = ws.Cells(i, 1).Value
                openPrice = ws.Cells(i, 3).Value
                totalVolume = ws.Cells(countTicker, 7).Value
            End If
        Next i
        'Make sure last line has all data printed
        ws.Cells(countTicker, 9).Value = ticker
        ws.Cells(countTicker, 10).Value = closePrice - openPrice
        ws.Cells(countTicker, 11).Value = ((closePrice - openPrice) / openPrice)
        ws.Cells(countTicker, 12).Value = totalVolume
        
        'Get number of rows that need formating
        Dim greatInc As Double
        Dim incTicker As String
        Dim greatDec As Double
        Dim decTicker As String
        Dim greatVolume As Double
        Dim volumeTicker As String
        Dim newRows As Integer
        
        greatInc = 0
        greatDec = 0
        greatVolume = 0
        newRows = ws.Application.CountA(ws.Range(ws.Cells(1, 9), ws.Cells(ws.Rows.Count, 9)))

        'format cell color for quarterly change while other data needed
        For j = 2 To newRows
            'get greatest increase, decrease, volume
            If ws.Cells(j, 11).Value > greatInc Then
                greatInc = ws.Cells(j, 11).Value
                incTicker = ws.Cells(j, 9).Value
            End If
            If ws.Cells(j, 11).Value < greatDec Then
                greatDec = ws.Cells(j, 11).Value
                decTicker = ws.Cells(j, 9).Value
            End If
            If ws.Cells(j, 12).Value > greatVolume Then
                greatVolume = ws.Cells(j, 12).Value
                volumeTicker = ws.Cells(j, 9).Value
            End If
            
            'change color format
            If ws.Cells(j, 10).Value > 0 Then
                ws.Cells(j, 10).Interior.Color = RGB(0, 255, 0)
            ElseIf ws.Cells(j, 10).Value < 0 Then
                ws.Cells(j, 10).Interior.Color = RGB(255, 0, 0)
            End If
        Next j
        
        'print greatest inc, dec, volume
        ws.Range("O2").Value = incTicker
        ws.Range("P2").Value = greatInc
        ws.Range("O3").Value = decTicker
        ws.Range("P3").Value = greatDec
        ws.Range("O4").Value = volumeTicker
        ws.Range("P4").Value = greatVolume
        
        'format percent change column
        ws.Range("K2:K" + CStr(newRows)).NumberFormat = "0.00%"
        ws.Range("P2").NumberFormat = "0.00%"
        ws.Range("P3").NumberFormat = "0.00%"
        
        'resize columns
        ws.Range("I1").EntireColumn.AutoFit
        ws.Range("J1").EntireColumn.AutoFit
        ws.Range("K1").EntireColumn.AutoFit
        ws.Range("L1").EntireColumn.AutoFit
        ws.Range("N1").EntireColumn.AutoFit
        ws.Range("O1").EntireColumn.AutoFit
        ws.Range("P1").EntireColumn.AutoFit
        
    Next ws
End Sub
'This is just to clear the worksheets
Sub clear()
    Dim ws As Worksheet

    For Each ws In Worksheets
        For i = 1 To 8
            ws.Range("I1").EntireColumn.Delete Shift:=xlShiftToLeft
        Next i
    Next ws
End Sub
