Sub stocks()

For Each ws In Worksheets

    'variables and default values
    Dim summary_table_row As Integer
    summary_table_row = 2
    Dim ticker As String
    Dim yearlychange As Double
    Dim percentchange As Double
    Dim closeprice As Double
    closeprice = 0
    Dim openprice As Double
    openprice = ws.Range("C2").Value
    Dim totalvolume As Double
    totalvolume = 0
    
    
    
    'headers and labels
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Value"
    
    'formating for percentage with two decimal places
    ws.Range("K2:K290").NumberFormat = "0.00%"
    
    'defining last row
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    'loop through columns to complete requested tasks
    For i = 2 To LastRow
    

      If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
      
        
        
        ticker = ws.Cells(i, 1).Value
        
        closeprice = ws.Cells(i, 6).Value
        
        totalvolume = totalvolume + ws.Cells(i, 7).Value
        yearlychange = closeprice - openprice
        
            If openprice = 0 Then
            percentchange = 0
            
            Else
            
            openprice = ws.Range("C2").Value
            percentchange = yearlychange / openprice
            
            End If
            
        
       'write data to cells
        ws.Range("I" & summary_table_row).Value = ticker
        ws.Range("L" & summary_table_row).Value = totalvolume
        ws.Range("J" & summary_table_row).Value = yearlychange
        ws.Range("K" & summary_table_row).Value = percentchange
        summary_table_row = summary_table_row + 1
        
        totalvolume = 0
    
        openprice = ws.Cells(i + 1, 3).Value
        
        Else
        totalvolume = totalvolume + ws.Cells(i, 7).Value
        
    End If

Next i


'conditional color formatting for positive and negative values
For j = 2 To LastRow

    If ws.Cells(j, 10).Value >= 0 Then
        
        ws.Cells(j, 10).Interior.ColorIndex = 4
       
    ElseIf ws.Cells(j, 10).Value < 0 Then
        
        ws.Cells(j, 10).Interior.ColorIndex = 3
    
    End If
    
Next j

'greatest increse, decrease, and total volume
LastRow = ws.Cells(Rows.Count, 11).End(xlUp).Row

ws.Range("Q2").NumberFormat = "%0.00"
ws.Range("Q3").NumberFormat = "%0.00"

For k = 2 To LastRow

    If ws.Cells(k + 1, 11).Value > ws.Range("Q2").Value Then
    ws.Range("Q2").Value = ws.Cells(k + 1, 11).Value
    ws.Range("P2").Value = ws.Cells(k + 1, 9).Value
    End If
     
    If ws.Cells(k + 1, 11).Value < ws.Range("Q3").Value Then
    ws.Range("Q3").Value = ws.Cells(k + 1, 11).Value
    ws.Range("P3").Value = ws.Cells(k + 1, 9).Value
    End If
   
    If ws.Cells(k + 1, 12).Value > ws.Range("Q4").Value Then
    ws.Range("Q4").Value = ws.Cells(k + 1, 12).Value
    ws.Range("P4").Value = ws.Cells(k + 1, 9).Value
    End If

Next k

ws.Columns("I:Q").AutoFit

Next ws


End Sub
