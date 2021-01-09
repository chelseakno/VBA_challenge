Sub NextTicker()
    
    Dim ws As Worksheet
    Dim openPrice As Double
    Dim closePrice As Double
    Dim yearlyChange As Double
    Dim tickerName As String
    Dim percentChange As Double
    Dim Volume As Double
        Volume = 0
    Dim Row As Double
        Row = 2
    Dim Column As Integer
        Column = 1
    Dim i As Long
        
    Const RED = 3
    Const GREEN = 4
        
        For Each ws In ThisWorkbook.Worksheets
        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
     
        
            For i = 2 To lastRow
                If Cells(i + 1, Column).Value <> Cells(i, Column).Value Then
                
                    tickerName = Cells(i, Column).Value
                    Cells(Row, Column + 8).Value = tickerName
                
                    closePrice = Cells(i, Column + 5).Value
                    openPrice = Cells(2, Column + 2).Value
                    yearlyChange = closePrice - openPrice
                    Cells(Row, Column + 9).Value = yearlyChange
                
                        If (openPrice = 0 And closePrice = 0) Then
                            percentChange = 0
                        ElseIf (openPrice = 0 And closePrice <> 0) Then
                        percentChange = 1
                        Else
                            percentChange = yearlyChange / openPrice
                            Cells(Row, Column + 10).Value = percentChange
                            Cells(Row, Column + 10).NumberFormat = "0.00%"
                        End If
                
                    Volume = Volume + Cells(i, Column + 6).Value
                    Cells(Row, Column + 11).Value = Volume
                    Row = Row + 1
                    openPrice = Cells(i + 1, Column + 2)
                    Volume = 0

                Else
                    Volume = Volume + Cells(i, Column + 6).Value
                End If
        
            Next i
  
  yearlyChangeLastRow = ws.Cells(Rows.Count, Column + 8).End(xlUp).Row
        
        For j = 2 To yearlyChangeLastRow
            If (Cells(j, Column + 9).Value > 0 Or Cells(j, Column + 9).Value = 0) Then
                Cells(j, Column + 9).Interior.ColorIndex = GREEN
            ElseIf Cells(j, Column + 9).Value < 0 Then
                Cells(j, Column + 9).Interior.ColorIndex = RED
            End If
            
        Next j
     
        Cells(1, Column + 15).Value = "Ticker"
        Cells(1, Column + 16).Value = "Value"
        Cells(2, Column + 14).Value = "Greatest % Increase"
        Cells(3, Column + 14).Value = "Greatest % Decrease"
        Cells(4, Column + 14).Value = "Greatest Total Volume"
        
        For Z = 2 To yearlyChangeLastRow
            If Cells(Z, Column + 10).Value = Application.WorksheetFunction.Max(ws.Range("K2:K" & yearlyChangeLastRow)) Then
                Cells(2, Column + 15).Value = Cells(Z, Column + 8).Value
                Cells(2, Column + 16).Value = Cells(Z, Column + 10).Value
                Cells(2, Column + 16).NumberFormat = "0.00%"
            ElseIf Cells(Z, Column + 10).Value = Application.WorksheetFunction.Min(ws.Range("K2:K" & yearlyChangeLastRow)) Then
                Cells(3, Column + 15).Value = Cells(Z, Column + 8).Value
                Cells(3, Column + 16).Value = Cells(Z, Column + 10).Value
                Cells(3, Column + 16).NumberFormat = "0.00%"
            ElseIf Cells(Z, Column + 11).Value = Application.WorksheetFunction.Max(ws.Range("L2:L" & yearlyChangeLastRow)) Then
                Cells(4, Column + 15).Value = Cells(Z, Column + 8).Value
                Cells(4, Column + 16).Value = Cells(Z, Column + 11).Value
            End If
        Next Z
        
    Next ws
        
End Sub

