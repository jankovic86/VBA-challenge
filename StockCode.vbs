Sub stocks()
'variables and default values
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


'headers and labels
Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Stock Volume"
Cells(1, 16).Value = "Ticker"
Cells(1, 17).Value = "Value"
Cells(2, 15).Value = "Greatest % Increase"
Cells(3, 15).Value = "Greatest % Decrease"
Cells(4, 15).Value = "Greatest Total Value"

'formating for percentage with two decimal places
Range("K2:K290").NumberFormat = "0.00%"

'defining last row
LastRow = Cells(Rows.Count, 1).End(xlUp).Row

'loop through columns to complete requested tasks
For i = 2 To LastRow
    

      If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
      
        
        
        ticker = Cells(i, 1).Value
        
        closeprice = Cells(i, 6).Value
        
        totalvolume = totalvolume + Cells(i, 7).Value
        yearlychange = closeprice - openprice
        
            If openprice = 0 Then
            percentchange = 0
            
            Else
            
            openprice = Range("C2").Value
            percentchange = yearlychange / openprice
            
            End If
            
        
       'write data to cells
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


'conditional color formatting for positive and negative values
For j = 2 To LastRow

    If Cells(j, 10).Value >= 0 Then
        
        Cells(j, 10).Interior.ColorIndex = 4
       
    ElseIf Cells(j, 10).Value < 0 Then
        
        Cells(j, 10).Interior.ColorIndex = 3
    
    End If
    
Next j

'greatest increse, decrease, and total volume
LastRow = Cells(Rows.Count, 11).End(xlUp).Row

Range("Q2").NumberFormat = "%0.00"
Range("Q3").NumberFormat = "%0.00"

For k = 2 To LastRow

    If Cells(k + 1, 11).Value > Range("Q2").Value Then
    Range("Q2").Value = Cells(k + 1, 11).Value
    Range("P2").Value = Cells(k + 1, 9).Value
    End If
     
    If Cells(k + 1, 11).Value < Range("Q3").Value Then
    Range("Q3").Value = Cells(k + 1, 11).Value
    Range("P3").Value = Cells(k + 1, 9).Value
    End If
   
    If Cells(k + 1, 12).Value > Range("Q4").Value Then
    Range("Q4").Value = Cells(k + 1, 12).Value
    Range("P4").Value = Cells(k + 1, 9).Value
    End If

Next k

Columns("I:Q").AutoFit

End Sub
