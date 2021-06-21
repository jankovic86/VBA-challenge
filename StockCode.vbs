Sub stocks()

Dim summary_table_row As Integer
summary_table_row = 2
Dim ticker As String
Dim yearchange As Long
Dim percentchange As Long
Dim totalvolume As Double
totalvolume = 0



Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Stock Volume"

LastRow = Cells(Rows.Count, 1).End(xlUp).Row

For i = 2 To LastRow
    

      If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
      
        ticker = Cells(i, 1).Value
        totalvolume = totalvolume + Cells(i, 7).Value
        
        Range("I" & summary_table_row).Value = ticker
        Range("L" & summary_table_row).Value = totalvolume
        
        summary_table_row = summary_table_row + 1
        
        totalvolume = 0
        
        Else
        
        totalvolume = totalvolume + Cells(i, 7).Value
        
        
    End If
    
    
Next i



End Sub
