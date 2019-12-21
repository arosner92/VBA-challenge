Attribute VB_Name = "Module1"
Sub stocks()
    Dim lastrow As Long
    Dim ticker As String
    Dim vol_count As Double
    Dim open_price As Single
    Dim close_price As Single
    Dim summary_row As Integer

    lastrow = Cells(Rows.Count, 1).End(xlUp).Row
    vol_count = 0
    summary_row = 2
    open_price = Cells(2, 3).Value
    
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
    
    For i = 2 To lastrow
        ticker = Cells(i, 1).Value

        If ticker <> Cells(i + 1, 1).Value Then
        
            vol_count = vol_count + Cells(i, 7).Value
            Cells(summary_row, 12).Value = vol_count
            vol_count = 0
            
            Cells(summary_row, 9).Value = ticker
            
            close_price = Cells(i, 6).Value

            Cells(summary_row, 10).Value = close_price - open_price
            Cells(summary_row, 10).NumberFormat = "0.00"
            If Cells(summary_row, 10).Value > 0 Then
                Cells(summary_row, 10).Interior.ColorIndex = 4
            Else
                Cells(summary_row, 10).Interior.ColorIndex = 3
            End If
            
            If (open_price = 0) And (close_price = 0) Then
                Cells(summary_row, 11).Value = 0
            ElseIf (open_price = 0) And (close_price <> 0) Then
                Cells(summary_row, 11).Value = "N/A"
            Else
                Cells(summary_row, 11).Value = (close_price - open_price) / open_price
            End If
            
            Cells(summary_row, 11).NumberFormat = "0.00%"
            
            open_price = Cells(i + 1, 3).Value
            summary_row = summary_row + 1
            
        Else
            vol_count = vol_count + Cells(i, 7).Value
            
        End If
    Next i
            
End Sub

