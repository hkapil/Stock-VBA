Sub TickerTotalsHard_2():


'Initialize variables that need to be set once before Worksheet Loop

Dim ws As Worksheet
Dim WorkSheetCount As Integer

WorkSheetCount = Sheets.Count ' Get TotalCount of Worksheets


'Loop through each Worksheet to Copy and Paste into Combined Data Tab
 
 For wsloop = 1 To WorkSheetCount

    'Reset variable for each worksheet

        Dim sht As Worksheet
        Dim LastRow As Long
        Dim TickerTotal As Variant
        Dim SummaryRow As Integer
        Dim Closing As Double
        Dim Opening As Double
        
        
        Dim MaxDecrease As Variant
        Dim MaxIncrease As Variant
        Dim MaxStockVolume As Variant
        Dim TickerNM_Inc As String
        Dim TickerNM_Dec As String
        Dim TickerNM_Stock As String
                
        
        ThisWorkbook.Worksheets(wsloop).Activate
        
        Set sht = ActiveSheet

        SummaryRow = 2
    'Ctrl + Shift + End to get last row of active sheet

        LastRow = sht.Cells(sht.Rows.Count, "A").End(xlUp).Row
    
        'Set Total Headers for active worksheet
        
        Cells(1, 9).Value = "Ticker"
        Cells(1, 10).Value = "Yearly Change"
        Cells(1, 11).Value = "Percentage Change"
        Cells(1, 12).Value = "Total Stock Volume"
    
    'Start loop for active worksheet to create summary totals for Tickers
    ' Get opening amount for firt Ticker in the active worksheet
    
    Opening = Cells(2, 3).Value
    MaxIncrease = Cells(2, 11).Value
    MaxDecrease = Cells(2, 11).Value
    MaxStockVolume = Cells(2, 12).Value
    TickerNM_Inc = Cells(2, 9).Value
    TickerNM_Dec = Cells(2, 9).Value
    TickerNM_Stock = Cells(2, 9).Value
    
    For i = 2 To LastRow
        
        TickerTotal = TickerTotal + Cells(i, 7).Value
        
        ' Ticker Change Level Summary
        
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            Closing = Cells(i, 6).Value
            Cells(SummaryRow, 9).Value = Cells(i, 1).Value
            
            Cells(SummaryRow, 10).Value = (Closing - Opening) 'Change in opening of the year Vs.Closing of the year
        
            'Calculate % change. Check for divide by Zero as when opening is Zero,divide by Zero occurs
        
            If Opening > 0 Then
                Cells(SummaryRow, 11).Value = (Cells(SummaryRow, 10).Value / Opening) * 100
            ElseIf Opening = 0 And Closing <> 0 Then
                Cells(SummaryRow, 11).Value = Cells(SummaryRow, 10).Value
            Else
                Cells(SummaryRow, 11).Value = 0
            End If
            
            Cells(SummaryRow, 12).Value = TickerTotal
            
          'Color Green for +ve, REd for -ve and leave white for Neutral or no change
          
            If Cells(SummaryRow, 10).Value < 0 Then
                Cells(SummaryRow, 10).Interior.ColorIndex = 3
            ElseIf Cells(SummaryRow, 10).Value > 0 Then
                Cells(SummaryRow, 10).Interior.ColorIndex = 4
            End If
            
            TickerTotal = 0
            Opening = Cells(i + 1, 3).Value
            
            SummaryRow = SummaryRow + 1
 
        End If

' Determine Sheet Level Summary of Greatest Increase, Decrease, Stock

        If (i <> LastRow) And Cells(i, 11).Value > MaxIncrease Then
            MaxIncrease = Cells(i, 11).Value
            TickerNM_Inc = Cells(i, 9).Value
        End If
        
        If (i <> LastRow) And Cells(i, 11).Value < MaxDecrease Then
            MaxDecrease = Cells(i, 11).Value
            TickerNM_Dec = Cells(i, 9).Value
        End If
                
        If i <> LastRow And Cells(i, 12).Value > MaxStockVolume Then
            MaxStockVolume = Cells(i, 12).Value
            TickerNM_Stock = Cells(i, 9).Value
        End If
        

    Next i
        
        
        'Set Headers & Values for Greatest %Increase, Greatest %Decrease and Greatest Stock Volume
        
        Cells(1, 16).Value = "Ticker"
        Cells(1, 17).Value = "Value"
        
        Cells(2, 15).Value = "Greatest %Increase"
        Cells(2, 16).Value = TickerNM_Inc
        Cells(2, 17).Value = MaxIncrease
        
        Cells(3, 15).Value = "Greatest %Decrease"
        Cells(3, 16).Value = TickerNM_Dec
        Cells(3, 17).Value = MaxDecrease
        
        Cells(4, 15).Value = "Greatest Total Volume"
        Cells(4, 16).Value = TickerNM_Stock
        Cells(4, 17).Value = MaxStockVolume
                
    Next wsloop

End Sub





