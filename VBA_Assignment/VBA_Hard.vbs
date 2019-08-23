Sub hard()

    Dim i, j, table_row As Integer
    Dim opening_price, closing_price, yearly_change, percent_change, volume, great_inc, great_dec, great_vol As Double
    Dim lastrow As Long
    Dim ticker, ticker_vol As String

    j = 2
    table_row = 2
    volume = 0
    great_inc = Cells(2, 11).Value
    ticker_inc = Cells(2, 9).Value
    great_dec = Cells(2, 11).Value
    ticker_dec = Cells(2, 9).Value
    great_vol = Cells(2, 12).Value
    ticker_vol = Cells(2, 9).Value

    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
    Range("O2").Value = "Greatest % increase"
    Range("O3").Value = "Greatest % Decrease"
    Range("O4").Value = "Greatest total volume"
    Range("P1").Value = "Ticker"
    Range("Q1").Value = "Value"
    
    lastrow = Cells(Rows.Count, 1).End(xlUp).Row
    
    For i = 2 To lastrow
    
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            ticker = Cells(i, 1).Value
            opening_price = Cells(j, 3).Value
            closing_price = Cells(i, 6).Value
            volume = volume + Cells(i, 7).Value
            
            If opening_price = 0 Then
                Range("J" & table_row).Value = 0
                Range("K" & table_row).Value = 0
                Range("L" & table_row).Value = 0
            Else
                yearly_change = closing_price - opening_price
                percent_change = ((closing_price / opening_price) - 1)
                Range("J" & table_row).Value = yearly_change
                Range("K" & table_row).Value = percent_change
                Range("L" & table_row).Value = volume
            End If
            
            Range("I" & table_row).Value = ticker
            
            table_row = table_row + 1
            j = i + 1
            volume = 0
            
        Else
            volume = volume + Cells(i, 7).Value
        End If
        
    Next i
    
    For i = 2 To lastrow
        
        If Cells(i, 10).Value > 0 Then
          Cells(i, 10).Interior.ColorIndex = 4
          
        Else
          Cells(i, 10).Interior.ColorIndex = 3
        End If
        
        If Cells(i + 1, 11).Value > great_inc Then
            great_inc = Cells(i + 1, 11).Value
            ticker_inc = Cells(i + 1, 9).Value
        
        ElseIf Cells(i + 1, 11).Value < great_dec Then
            great_dec = Cells(i + 1, 11).Value
            ticker_dec = Cells(i + 1, 9).Value
        End If
        
        If Cells(i + 1, 12).Value > great_vol Then
            great_vol = Cells(i + 1, 12).Value
            ticker_vol = Cells(i + 1, 9).Value
        End If
            
    Next i

    Columns("K").NumberFormat = "0.00%"
    Range("P2").Value = ticker_inc
    Range("Q2").Value = great_inc
    Range("P3").Value = ticker_dec
    Range("Q3").Value = great_dec
    Range("P4").Value = ticker_vol
    Range("Q4").Value = great_vol
    Range("Q2").NumberFormat = "0.00%"
    Range("Q3").NumberFormat = "0.00%"
    
End Sub

