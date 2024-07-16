Attribute VB_Name = "Module1"
Sub Module2()
Dim ws As Worksheet
Dim lastRow As Long
Dim ticker As String
Dim closePrice As Double
Dim openPrice As Double
Dim quarterStartDate As Date
Dim Volume As Double
Dim outputRow As Long
Dim cell As Range

For Each ws In ThisWorkbook.Worksheets
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Quarterly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    Volume = 0
    outputRow = 2
    ticker = ws.Cells(2, 1).Value
            openPrice = ws.Cells(2, 3).Value
    For i = 2 To lastRow
    Volume = Volume + ws.Cells(i, 7).Value
        If (ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value) Then
    
        
            closePrice = ws.Cells(i, 6).Value
            ws.Cells(outputRow, 9).Value = ticker
            quarterChange = closePrice - openPrice
            ws.Cells(outputRow, 10).Value = quarterChange
            ws.Cells(outputRow, 11).Value = quarterChange / openPrice
            
            ws.Cells(outputRow, 12).Value = Volume
            ws.Cells(outputRow, 11).NumberFormat = "0.00%"
            
            
            Select Case quarterChange
                Case Is > 0
                    ws.Cells(outputRow, 10).Interior.ColorIndex = 4 ' Green for > 100
            
                Case Is < 0
                     ws.Cells(outputRow, 10).Interior.ColorIndex = 3  ' Red for < 50
                Case Else
                     ws.Cells(outputRow, 10).Interior.ColorIndex = 0 ' No color
            
            
            End Select
            
            
            outputRow = outputRow + 1
           
            
            ticker = ws.Cells(i, 1).Value
            openPrice = ws.Cells(i + 1, 3).Value
            Volume = 0
            End If
            
            
    
            
            
            Next i
            Next
        

        
    End Sub

