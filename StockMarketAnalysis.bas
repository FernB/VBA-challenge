Attribute VB_Name = "Module1"
Sub StockMarketAnalyses()

Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual


Dim i, j, k, LstRow, lstRows As Long
Dim Ticker, TickerInc, TickerDec, TickerVol As String
Dim Opening, Closing, TotalVolume, MaxInc, MaxDec, MaxTotalVol As Double


For Each ws In Worksheets

'Reset counters at next ws
LstRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
j = 2
Ticker = vbNullString: TickerInc = vbNullString: TickerDec = vbNullString: TickerVol = vbNullString
Opening = 0: Closing = 0: TotalVolume = 0: MaxInc = 0: MaxDec = 0: MaxTotalVol = 0

'Loop to find each data summary per ticker
For i = 2 To LstRow

    If ws.Cells(i, 1).Value <> ws.Cells((i - 1), 1).Value And i = 2 Then
        Ticker = ws.Cells(i, 1).Value
        Opening = ws.Cells(i, 3).Value
        TotalVolume = ws.Cells(i, 7).Value
        
    ElseIf ws.Cells(i, 1).Value <> ws.Cells((i - 1), 1).Value Then
        Ticker = ws.Cells(i, 1).Value
        Opening = ws.Cells(i, 3).Value
        TotalVolume = ws.Cells(i, 7).Value
        

    ElseIf (ws.Cells(i, 1).Value = ws.Cells((i - 1), 1).Value) And (ws.Cells((i + 1), 1).Value = ws.Cells(i, 1).Value) Then
        TotalVolume = TotalVolume + ws.Cells(i, 7).Value

    ElseIf (ws.Cells(i, 1).Value = ws.Cells((i - 1), 1).Value) And (ws.Cells((i + 1), 1).Value <> ws.Cells(i, 1).Value) Then
        TotalVolume = TotalVolume + ws.Cells(i, 7).Value
        Closing = ws.Cells(i, 6).Value

        ws.Cells(j, 9).Value = Ticker
        ws.Cells(j, 10).Value = Closing - Opening
        If ws.Cells(j, 10).Value > 0 Then
            ws.Cells(j, 10).Interior.Color = RGB(0, 255, 0)
        ElseIf ws.Cells(j, 10).Value < 0 Then
            ws.Cells(j, 10).Interior.Color = RGB(255, 0, 0)
        End If
        ws.Cells(j, 12).Value = TotalVolume
        If Opening > 0 Then
           ws.Cells(j, 11).Value = ((Closing - Opening) / Opening)
           ws.Cells(j, 11).NumberFormat = "0.00%"
        Else
            ws.Cells(j, 11).Value = 0
        End If
        j = j + 1
    End If
    
    
Next i

'Adds headers
ws.Range("I1:L1").Value = Array("Ticker", "Yearly Change", "Percentage Change", "Total Stock Volume")

'Loop to find Max and Min Stats


lstRows = ws.Cells(Rows.Count, 9).End(xlUp).Row

For k = 2 To lstRows
    If ws.Cells(k, 11).Value > MaxInc Then
    MaxInc = ws.Cells(k, 11).Value
    TickerInc = ws.Cells(k, 9).Value
        If ws.Cells(k, 12) > MaxTotalVol Then
        MaxTotalVol = ws.Cells(k, 12).Value
        TickerVol = ws.Cells(k, 9).Value
        End If
    ElseIf ws.Cells(k, 11).Value < MaxDec Then
    MaxDec = ws.Cells(k, 11).Value
    TickerDec = ws.Cells(k, 9).Value
        If ws.Cells(k, 12).Value > MaxTotalVol Then
        MaxTotalVol = ws.Cells(k, 12).Value
        TickerVol = ws.Cells(k, 9).Value
        End If
    ElseIf ws.Cells(k, 12).Value > MaxTotalVol Then
    MaxTotalVol = ws.Cells(k, 12).Value
    TickerVol = ws.Cells(k, 9).Value
    End If
    
Next k


'Fills out Max and Min Table
ws.Range("O2").Value = "Greatest % Increase"
ws.Range("O3").Value = "Greatest % Decrease"
ws.Range("O4").Value = "Greatest Total Volume"
ws.Range("P1").Value = "Ticker"
ws.Range("P2").Value = TickerInc
ws.Range("P3").Value = TickerDec
ws.Range("P4").Value = TickerVol
ws.Range("Q1").Value = "Value"
ws.Range("Q2").Value = MaxInc
ws.Range("Q3").Value = MaxDec
ws.Range("Q2:Q3").NumberFormat = "0.00%"
ws.Range("Q4").Value = MaxTotalVol
ws.Range("Q4").NumberFormat = "0.00E+00"

'Formats Columns
ws.Range("A:Q").Columns.AutoFit


Next ws
Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic

MsgBox ("Analysis Completed")

End Sub

