Attribute VB_Name = "Module1"
Sub tickeranalysis()

    For Each WS In Worksheets

        WS.Activate
        
        Range("I1").Value = "Ticker"
        Range("J1").Value = "Yearly Change"
        Range("K1").Value = "Percent Change"
        Range("L1").Value = "Total Stock Volume"
        
        Columns("I:L").AutoFit
        
        total_volume = 0
        
        row_count = Cells(Rows.Count, "A").End(xlUp).Row

        
        For i = 2 To row_count
        
            If Cells(i, "A").Value = Cells(i + 1, "A").Value Then
                total_volume = total_volume + Cells(i, "G").Value
                
            Else
               total_volume = total_volume + Cells(i, "G").Value
               
               
               total_volume = 0
            End If
        

        Next i


    Next WS
End Sub
