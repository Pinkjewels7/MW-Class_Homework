Sub Stock_Easy():

Dim ticker_name As String
Dim vol_total As Double
Dim lastrow, sum_row As Integer

ticker_total = 0
lastrow = Cells(Rows.Count, 1).End(xlUp).Row
sum_row = 2

    For i = 2 To lastrow

        If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
        
            ticker_name = Cells(i, 1).Value
            vol_total = vol_total + Cells(i, 7).Value
            
            Range("i" & sum_row).Value = ticker_name
            Range("j" & sum_row).Value = vol_total
            sum_row = sum_row + 1

            ticker_total = 0

        Else

            vol_total = vol_total + Cells(i, 7).Value
 
        End If

    Next i

    'formatting
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Tot. Vol."
    
    Range("I1:J1").Select
    Selection.Font.Bold = True
    Selection.Interior.Color = 65535
    Selection.HorizontalAlignment = xlCenter
    
    Columns("J:J").Select
    Selection.NumberFormat = "#,##0_);[Red](#,##0)"
    
    Range("A1").Select

End Sub

