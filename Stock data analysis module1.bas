Attribute VB_Name = "Module1"
Sub StockMarket():

Dim ws As Worksheet
Dim yearly_change As Double
Dim percent_change As Double
Dim total_vol As Double

Dim ticker_col As Integer
ticker_col = 1

Dim date_col As Integer
date_col = 2
Dim open_col As Integer
open_col = 3
Dim open_price As Double
Dim close_col As Integer
close_col = 6
Dim close_price As Double
Dim vol_col As Integer
vol_col = 7
Dim vol As Integer

Dim symbol As String
Dim lastrow As Long
Dim i As Long
Dim symbol_col As Integer
symbol_col = 9
Dim yealy_change_col As Integer
yearly_change_col = 10
Dim percent_change_col As Integer
percent_change_col = 11
Dim total_vol_col As Integer
total_vol_col = 12
Dim output_row As Integer


For Each ws In Worksheets
    Sheets(ws.Name).Select

    output_row = 2
    total_vol = 0
    
    'Get the number of rows in the current worksheet
    lastrow = Cells(Rows.Count, 1).End(xlUp).Row
    
    'Set the open price for the first stock in the current worksheet
    open_price = Cells(2, open_col).Value
    
    'set the cell headings for the output columns
    Cells(1, symbol_col).Value = "Ticker"
    Cells(1, yearly_change_col).Value = "Yearly Change"
    Cells(1, percent_change_col).Value = "Percent Change"
    Cells(1, total_vol_col).Value = "Total Volume"
        
        For i = 2 To lastrow
        
            'Check to see if the ticker symbols of the current row and 1 after it do not match
            If Cells(i + 1, ticker_col).Value <> Cells(i, ticker_col).Value Then
            
            'get the ticker symbol of the stock you're on
            symbol = Cells(i, 1).Value
            
            'get the year end close price the last close price in column F.
            close_price = Cells(i, close_col).Value
            
            'calculate the yearly change by subtracting the year end close price by the open price at the beginning of the year.
            yearly_change = close_price - open_price
            
            'calculate the percent change for the year. Added if statement for divide by 0 scenerios.
            If open_price = 0 And close_price = 0 Then
                percent_change = 0
            ElseIf open_price = 0 And close_price <> 0 Then
                percent_change = yearly_change
            Else
            percent_change = (close_price - open_price) / open_price
            End If
            
            'calculate the total volume by adding the value stored in total_vol and the current rows total volume in column G.
            total_vol = total_vol + Cells(i, vol_col).Value
            
            'place all the output values in columns H,I,J,K next to the table.
            Range("I" & output_row).Value = symbol
            Range("J" & output_row).Value = yearly_change
            'Add the percent change column with conditional formatting to make the cell green if +, red if -, and yellow if 0.
            Range("K" & output_row).Value = percent_change
            If percent_change > 0 Then
                    Range("K" & output_row).Interior.Color = vbGreen
                ElseIf percent_change < 0 Then
                    Range("K" & output_row).Interior.Color = vbRed
                Else
                    Range("K" & output_row).Interior.Color = vbYellow
                End If
            'format the percent change column as a percent
            Columns(percent_change_col).NumberFormat = "0.00%"
            
            Range("L" & output_row).Value = total_vol
            
            'add one to the output row so it continues to add each ticker in their own row.
            output_row = output_row + 1
            
            'set the start of the year open price for the next ticker by storing the first open price for that symbol
            open_price = Cells(i + 1, open_col).Value
            
            'reset total volume and close price to 0 and output row to 2
            total_vol = 0
            close_price = 0
            
            
            Else
            
            'if the ticker symbols match, continue to add the total volume for that symbol
            total_vol = total_vol + Cells(i, vol_col).Value
            
            
            End If
        
        Next i
        
        


Next ws

End Sub

