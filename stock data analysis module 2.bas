Attribute VB_Name = "Module2"
Sub greatestmovers():

Dim greatest_increase_col As Integer
Dim greatest_increase As Double
Dim greatest_increase_symbol As String
Dim greatest_decrease_col As Integer
Dim greatest_decrease As Double
Dim greatest_decrease_symbol As String
Dim greatest_tot_vol_col As Integer
Dim greatest_tot_vol As Double
Dim greatest_vol_symbol As String
Dim lastrow As Long
Dim i As Long
Dim j As Integer


'iterate through each worksheet in the workbook.
For Each ws In Worksheets
    
    'select the current sheet.
    Sheets(ws.Name).Select
    
    'Get the last row of the output table of the current sheet.
    lastrow = Cells(Rows.Count, 9).End(xlUp).Row
        
        'Enter in the column and row names of the output rows and columns
        Cells(2, 15).Value = "Greatest Increase"
        Cells(3, 15).Value = "Greatest Decrease"
        Cells(4, 15).Value = "Greatest Total Volume"
        Cells(1, 16).Value = "Ticker"
        Cells(1, 17).Value = "value"
        
        'set the greatest increase, greatest decrease, and greatest total volume variables to 0
        greatest_increase = 0
        greatest_decrease = 0
        greatest_tot_vol = 0
            
            'start the for loop with the columns so that both the greatest percent change values and the greatest total volume are captured.
            For j = 11 To 12
                
                'nested for loop so that the code will loop through all of the rows in the percent column first then every row in the total volume column.
                For i = 2 To lastrow
            
                    'If statement to compare the current cell in the percent change column to the greatest increase value.
                    'If the current cell is greater than make that the new greatest increase value and store the corresponding symbol. If it's not then do nothing and continue to the next cell.
                    If Cells(i, 11).Value > greatest_increase Then
                    greatest_increase = Cells(i, 11).Value
                    greatest_increase_symbol = Cells(i, 9).Value
                    Else
                    End If
                        
                    'If statement to compare the current cell in the percent change column to the greatest decrease value.
                    'If the current cell is less than the greatest decrease value then make it the new greatest decrease value and store the corresponding symbol. If it's not then do nothing and continue to the next cell.
                    If Cells(i, 11).Value < greatest_decrease Then
                    greatest_decrease = Cells(i, 11).Value
                    greatest_decrease_symbol = Cells(i, 9).Value
                    Else
                    End If
                    
                    'If statement to compare the current cell in the total volume column to the greatest total volumn value.
                    'If the current cell in the greatest total volume column is greater than the greatest total volue value, then make it the new value and store the corresponding symbol. If not, then do nothing and move to the next cell.
                    If Cells(i, 12).Value > greatest_tot_vol Then
                    greatest_tot_vol = Cells(i, 12).Value
                    greatest_vol_symbol = Cells(i, 9).Value
                    Else
                    End If
    
        
                Next i
            Next j
                
                'Take the values that we stored from out If statements and enter them into the output rows. Format the greatest percent change cells as percents.
                Cells(2, 16).Value = greatest_increase_symbol
                Cells(2, 17).Value = greatest_increase
                Cells(2, 17).NumberFormat = "0.00%"
                Cells(3, 16).Value = greatest_decrease_symbol
                Cells(3, 17).Value = greatest_decrease
                Cells(3, 17).NumberFormat = "0.00%"
                Cells(4, 16).Value = greatest_vol_symbol
                Cells(4, 17).Value = greatest_tot_vol
        Next ws
        
End Sub

