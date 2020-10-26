Attribute VB_Name = "Module2"
Sub StockAnalyzer()
    
    Dim dataArr As Variant
    Dim total_volume As LongLong
    Dim yearly_changes As Double
    Dim percent_changes As Double
    Dim summary_row As Long
    Dim ws As Worksheet
    Dim open_price As Double
    Dim close_price As Double
    
    
    
    For Each ws In Worksheets
    
        last_row = ws.Cells(Rows.Count, 1).End(xlUp).row
        
         'Add title to column J, K, L, M
        ws.Range("J1").Value = "Ticker"
        ws.Range("K1").Value = "Yearly change"
        ws.Range("L1").Value = "Percent Change"
        ws.Range("M1").Value = "Total stock volume"
        
        'Adjust cell size
        ws.Range("J1:M" & last_row).Columns.AutoFit
        
        ws.Range("L:L").NumberFormat = "0.00%"
        
    
        dataArr = ws.Range("A2:G" & last_row).Value
        total_volume = 0
        yearly_changes = 0
        percent_changes = 0
        open_price = dataArr(2, 3)
        close_price = 0
        summary_row = 2
        
        For r_index = LBound(dataArr, 1) To UBound(dataArr, 1) - 1
            total_volume = total_volume + dataArr(r_index, 7)
            
            If dataArr(r_index, 1) <> dataArr(r_index + 1, 1) Then
                ws.Cells(summary_row, "J").Value = dataArr(r_index, 1)
                
                close_price = dataArr(r_index, 6)
                ' yearly changes
                yearly_changes = close_price - open_price
                ws.Cells(summary_row, "K").Value = yearly_changes
                
                If yearly_changes <= 0 Then
                    ws.Cells(summary_row, "K").Interior.ColorIndex = 3
                Else
                    ws.Cells(summary_row, "K").Interior.ColorIndex = 4
                End If
                
                
                'persent changes
                If open_price = close_price Then
                    persent_changes = 0
                ElseIf open_price = 0 And close_price > 0 Then
                    percent_changes = ((close_price + 1) - (open_price + 1)) / (open_price + 1)
                ElseIf open_price > 0 And close_price = 0 Then
                    percent_changes = 100
                Else
                    percent_changes = (close_price - open_price) / open_price
                    
                End If
                
                ws.Cells(summary_row, "L").Value = percent_changes
                ws.Cells(summary_row, "M").Value = total_volume
                
                open_price = dataArr(r_index + 1, 3)
                yearly_changes = 0
                percent_changes = 0
                total_volume = 0
                summary_row = summary_row + 1
            End If
        Next r_index
        
        'Add title to columns
        ws.Range("P2").Value = "Greatest % increase"
        ws.Range("P3").Value = "Greatest % decrease"
        ws.Range("P4").Value = "Greatest total volume"
        ws.Range("Q1").Value = "Ticker"
        ws.Range("R1").Value = "Total"
        ws.Range("R2:R3").NumberFormat = "0.00%"
        
        'Adjust cell size
        ws.Range("P1:R" & last_row).Columns.AutoFit
        
        ' assign first value as a greates increase
        ws.Range("R2").Value = ws.Range("L2").Value
        ws.Range("Q2").Value = ws.Range("J2").Value
        
        ' assign first value as a greatest decrease
        ws.Range("R3").Value = ws.Range("L2").Value
        ws.Range("Q3").Value = ws.Range("J2").Value
        
        'assign first value as a greatest total volume
        ws.Range("R4").Value = ws.Range("M2").Value
        ws.Range("Q4").Value = ws.Range("J2").Value

        
        For row = 2 To last_row
        ' find greatest increase
            If ws.Range("R2").Value < ws.Cells(row, "L").Value Then

                ws.Range("R2").Value = ws.Cells(row, "L").Value
                ws.Range("Q2").Value = ws.Cells(row, "J").Value
                
            End If
        ' find greatest decrease
            If ws.Range("R3").Value > ws.Cells(row, "L") Then

                ws.Range("R3").Value = ws.Cells(row, "L").Value
                ws.Range("Q3").Value = ws.Cells(row, "J").Value
            End If
        ' find max total volume
            If ws.Range("R4").Value < ws.Cells(row, "M").Value Then
                ws.Range("R4").Value = ws.Cells(row, "M").Value
                ws.Range("Q4").Value = ws.Cells(row, "J").Value
            End If
        
        Next row
              
    
    Next ws

End Sub
