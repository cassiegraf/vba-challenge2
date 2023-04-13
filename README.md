# vba-challenge2
module 2 assignment

    Sub StockAssignment()

    For Each ws In ThisWorkbook.Worksheets
    ws.Activate


    ws.Range("I1").Value = "ticker symbol"
    ws.Range("J1").Value = "total stock volume"
    ws.Range("k1").Value = "yearly change ($)"
    ws.Range("L1").Value = "percent change"
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    
    

  
    Dim Ticker As String

    Dim VolumeTotal As LongLong
    VolumeTotal = 0

    Dim OpenPrice As Double
    Dim ClosePrice As Double
    Dim YearlyChange As Double

    Dim StockTable As Integer
    StockTable = 2 'for row 2

    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row



    For i = 2 To LastRow

    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        Ticker = ws.Cells(i, 1).Value
        
        VolumeTotal = VolumeTotal + ws.Cells(i, 7).Value
        
        ws.Range("I" & StockTable).Value = Ticker
        
        ws.Range("J" & StockTable).Value = VolumeTotal
                
        VolumeTotal = 0
        
        ClosePrice = ws.Cells(i, 6).Value
        
        YearlyChange = ClosePrice - OpenPrice
        
        ws.Range("K" & StockTable).Value = YearlyChange
        
        PercentChange = ((ClosePrice - OpenPrice) / OpenPrice)
                       
        ws.Range("l" & StockTable).Value = PercentChange
        ws.Range("l" & StockTable).NumberFormat = "0.00%"
        
        
        
        StockTable = StockTable + 1
        
    ElseIf ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then
        OpenPrice = ws.Cells(i, 3).Value
        
            
                      
    Else
        VolumeTotal = VolumeTotal + ws.Cells(i, 7).Value
                     
        
    End If
    
    If ws.Cells(i, 11).Value > 0 Then
        ws.Cells(i, 11).Interior.ColorIndex = 4
        
        ElseIf ws.Cells(i, 11).Value < 0 Then
        ws.Cells(i, 11).Interior.ColorIndex = 3
        
    End If
    
    If ws.Cells(i, 12).Value > 0 Then
        ws.Cells(i, 12).Interior.ColorIndex = 4
        
        ElseIf ws.Cells(i, 12).Value < 0 Then
        ws.Cells(i, 12).Interior.ColorIndex = 3
        
    End If
    
        
    
    Next i


    ws.Range("Q2").Value = "%" & WorksheetFunction.Max(ws.Range("L2:L" & LastRow)) * 100
    ws.Range("Q3").Value = "%" & WorksheetFunction.Min(ws.Range("L2:L" & LastRow)) * 100
    ws.Range("Q4").Value = WorksheetFunction.Max(ws.Range("J2:J" & LastRow))





    Next ws



    End Sub



