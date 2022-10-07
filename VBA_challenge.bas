Attribute VB_Name = "Module1"
Sub mutiple_year_stock_data()


Dim ticket As String
Dim total_volume As Double
Dim yearly_change As Double
Dim lr As Long
Dim percent_change As Double
Dim summary_table As Long
Dim i As Long
Dim start As Double


Range("I1").Value = "Ticker"
Range("J1").Value = "Yearly Change"
Range("k1").Value = "Percent Change"
Range("L1").Value = "Total Stock Value"

total_volume = 0
yearly_change = 0
summary_table = 2
start = 2
lr = Cells(Rows.Count, "A").End(xlUp).Row



            For i = 2 To lr
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        total_volume = total_volume + Cells(i, 7).Value
        If total_volume = 0 Then
            Range("I" & summary_table).Value = Cells(i, 1).Value
            Range("J" & summary_table).Value = 0
            Range("K" & summary_table).Value = 0
            Range("L" & summary_table).Value = 0
        Else
            If Cells(start, 3) = 0 Then
                For Value = start To i
                    If Cells(Value, 3).Value <> 0 Then
                        start = Value
                Exit For
        End If
        Next Value
        End If
        yearly_change = Cells(i, 6) - Cells(start, 3)
        percent_change = yearly_change / Cells(start, 3)
        start = i + 1
        Range("I" & summary_table).Value = Cells(i, 1).Value
        Range("L" & summary_table).Value = total_volume
        Range("J" & summary_table).Value = yearly_change
        Range("K" & summary_table).Value = percent_change
        Range("J" & summary_table).NumberFormat = "0.00"
        Range("K" & summary_table).NumberFormat = "0.00%"
        If percent_change > 0 Then
            Range("J" & summary_table).Interior.ColorIndex = 4
        End If
        If percent_change < 0 Then
            Range("J" & summary_table).Interior.ColorIndex = 3
        End If
        End If
        total_volume = 0
        summary_table = summary_table + 1
    Else
    
        total_volume = total_volume + Cells(i, 7).Value
    End If
    
    Next i
        
        

End Sub
