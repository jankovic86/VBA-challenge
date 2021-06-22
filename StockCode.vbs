Sub stocks()

Dim summary_table_row As Integer
summary_table_row = 2
Dim ticker As String
Dim yearlychange As Double
Dim percentchange As Double
Dim closeprice As Double
closeprice = 0

Dim openprice As Double
openprice = Range("C2").Value




Dim totalvolume As Double
totalvolume = 0



Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Stock Volume"

Range("K2:K290").NumberFormat = "0.00%"


LastRow = Cells(Rows.Count, 1).End(xlUp).Row


For i = 2 To LastRow
    

      If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
      
        
        
        ticker = Cells(i, 1).Value
        
        closeprice = Cells(i, 6).Value
        
        totalvolume = totalvolume + Cells(i, 7).Value
        yearlychange = closeprice - openprice
        percentchange = yearlychange / openprice
        
        
        
        
        Range("I" & summary_table_row).Value = ticker
        Range("L" & summary_table_row).Value = totalvolume
        Range("J" & summary_table_row).Value = yearlychange
        Range("K" & summary_table_row).Value = percentchange
        summary_table_row = summary_table_row + 1
        
        totalvolume = 0
    
        openprice = Cells(i + 1, 3).Value
        
       
        
        
        Else
        totalvolume = totalvolume + Cells(i, 7).Value
        
    End If

Next i

For j = 2 To LastRow

    If Cells(j, 10).Value > 0 Then
        
        Cells(j, 10).Interior.ColorIndex = 4
       
    ElseIf Cells(j, 10).Value < 0 Then
        
        Cells(j, 10).Interior.ColorIndex = 3
    
       
    End If

Next j


End Sub
