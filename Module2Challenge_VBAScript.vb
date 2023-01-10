'declare variables, applied to every subroutine in this workbook
Dim GreatestPercentDecrease As Double
Dim GreatestPercentDecreaseTicker As String
Dim GreatestPercentIncrease As Double
Dim GreatestPercentIncreaseTicker As String
Dim GreatestTotalVolume As LongLong
Dim GreatestTotalVolumeTicker As String

Dim lr1 As Long
Dim lr2 As Long
Dim OpenPrice As Double
Dim PercentChange As Double
Dim TickerCode As String
Dim SummaryRow As Integer
Dim TotalStockVolume As LongLong
Dim ws As Worksheet
Dim YearlyChange As Double

Sub Module2ChallengeFinal()
    
    For Each ws In ThisWorkbook.Worksheets
               
        'first table labels
        ws.Cells(1, 9) = "Ticker"
        ws.Cells(1, 10) = "Yearly Change"
        ws.Cells(1, 11) = "Percent Change"
        ws.Cells(1, 12) = "Total Stock Volume"
    
        'second table labels
        ws.Cells(1, 15) = "Ticker"
        ws.Cells(1, 16) = "Value"
        ws.Cells(2, 14) = "Greatest % Increase"
        ws.Cells(3, 14) = "Greatest % Decrease"
        ws.Cells(4, 14) = "Greatest Total Volume"
        
        'find the last cell (ie control A)
        lr1 = ws.Cells(Rows.Count, 1).End(xlUp).Row
        lr2 = ws.Cells(Rows.Count, 9).End(xlUp).Row
        
        'set variable values
        YearlyChange = 0
        OpenPrice = ws.Cells(2, 3).Value
        PercentChange = 0
        TotalStockVolume = 0
        SummaryRow = 2
        
    'loop through provided data except first row
    For i = 2 To lr1
    
        'find new ticker relative to previous cell selection
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
            TickerCode = Cells(i, 1).Value
            
            'total stock volume for each ticker
            TotalStockVolume = TotalStockVolume + ws.Cells(i, 7).Value
            
            'closing price - open
            YearlyChange = ws.Cells(i, 6).Value - OpenPrice
            
                    'calculate percent change, avoid dividing by 0
                    If OpenPrice = 0 Then
            
                        PercentChange = 0
            
                        Else
            
                            PercentChange = YearlyChange / OpenPrice
            
                    End If
                
                'print tickercode in summary table
                ws.Range("I" & SummaryRow).Value = TickerCode
            
                'print yearly change in summary table
                ws.Range("J" & SummaryRow).Value = YearlyChange
        
            
                    'conditional formatting for negative and then positive
                    If YearlyChange < 0 Then
        
                        ws.Range("J" & SummaryRow).Interior.Color = rgbTomato
        
                        ElseIf YearlyChange >= 0 Then
        
                            ws.Range("J" & SummaryRow).Interior.Color = rgbDarkSeaGreen
            
                        Else
        
                            ws.Range("J" & SummaryRow).Interior.Color = False
            
                    End If
        
                'print percent change in summary table
                ws.Range("K" & SummaryRow).Value = Format(PercentChange, "percent")
            
                'print total stock volume
                ws.Range("L" & SummaryRow).Value = TotalStockVolume
            
                'go to next row in summary table
                SummaryRow = SummaryRow + 1
                            
                'reset total stock volume to 0
                TotalStockVolume = 0
            
                OpenPrice = ws.Cells(i + 1, 3).Value
            
                Else
                    'add volume
                    TotalStockVolume = TotalStockVolume + ws.Cells(i, 7).Value
        
            End If
        
    Next i
    
    'autofit table columns
    ws.Range("I1:L1").EntireColumn.AutoFit
    
    'use the max and min function to find value in the respective column
    GreatestPercentIncrease = WorksheetFunction.Max((Range(ws.Cells(2, 11), ws.Cells(lr2, 11))))
    GreatestPercentDecrease = WorksheetFunction.Min((Range(ws.Cells(2, 11), ws.Cells(lr2, 11))))
    GreatestTotalVolume = WorksheetFunction.Max((Range(ws.Cells(2, 12), ws.Cells(lr2, 12))))
    
    'loop through provided data except first row
    For i = 2 To lr2
    
        'write value in new summary table
        If GreatestTotalVolumeTicker = ws.Cells(i, 12).Value Then
                
            Range("P3").Value = ws.Cells(i, 9).Value
        
            ElseIf GreatestPercentDecreaseTicker = ws.Cells(i, 11).Value Then
            
                Range("P2").Value = ws.Cells(i, 9).Value
                
            ElseIf GreatestPercentIncreaseTicker = ws.Cells(i, 11).Value Then
            
                Range("O2").Value = ws.Cells(i, 9).Value
                
        End If
        
    Next i
            
    
    'write values in last summary table
    ws.Cells(2, 15).Value = GreatestPercentIncreaseTicker
    ws.Cells(2, 16).Value = Format(GreatestPercentIncrease, "percent")
    ws.Cells(3, 15).Value = GreatestPercentDecreaseTicker
    ws.Cells(3, 16).Value = Format(GreatestPercentDecrease, "percent")
    ws.Cells(4, 15).Value = GreatestTotalVolumeTicker
    ws.Cells(4, 16).Value = GreatestTotalVolume
                    
   'reset values
   GreatestPercentIncreaseTicker = ""
   GreatestPercentDecreaseTicker = ""
   GreatestTotalVolumeTicker = ""
    
   'autofit table columns
    ws.Range("N1:P1").EntireColumn.AutoFit
    
   Next ws


End Sub