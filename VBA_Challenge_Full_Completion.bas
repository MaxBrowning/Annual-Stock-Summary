Attribute VB_Name = "VBA_Challenge"
Sub VBA_Challenge()

For Each ws In Worksheets

'Define all variables
Dim First_Open_Value As Double
Dim Last_Close_Value As Double
Dim Yearly_Change As Double
Dim Total_Stock_Volume As Double
Dim max As Double
Dim min As Double

'Set new headers
ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "Total Stock Volume"
ws.Cells(2, 14).Value = "Greatest % Increase"
ws.Cells(3, 14).Value = "Greatest % Decrease"
ws.Cells(4, 14).Value = "Greatest Total Volume"
ws.Cells(1, 15).Value = "Ticker"
ws.Cells(1, 16).Value = "Value"

'Set the row two's values
ws.Cells(2, 9).Value = ws.Cells(2, 1).Value
First_Open_Value = ws.Cells(2, 3).Value
Total_Stock_Volume = ws.Cells(2, 7).Value

'j will define the row number for column Ticker
j = 2

'i will define the row number for column <ticker>
For i = 2 To ws.Cells(Rows.Count, "A").End(xlUp).Row
    
    'Skip lines where First_Open_Values = 0
    If ws.Cells(i, 3).Value = 0 And ws.Cells(i, 6).Value = 0 Then
    
    ElseIf ws.Cells(i, 1).Value <> ws.Cells(j, 9).Value Then
    
        'Go to the next row.
        j = j + 1
        
        'Print new ticker symbol
        ws.Cells(j, 9).Value = ws.Cells(i, 1).Value
        
        'Save last close value
        Last_Close_Value = ws.Cells(i - 1, 6).Value
        
        'Calculate yearly change
        ws.Cells(j - 1, 10).Value = Last_Close_Value - First_Open_Value
        Yearly_Change = ws.Cells(j - 1, 10).Value
        
        'Print the previous brand's percent change
        ws.Cells(j - 1, 11).Value = (Yearly_Change / First_Open_Value)
        
        'Print the previous brand's total stock volume
        ws.Cells(j - 1, 12).Value = Total_Stock_Volume
        
        'Reset Total_Stock_Volume starting value for new brand
        Total_Stock_Volume = ws.Cells(i, 7).Value
        
        'Reset First_Open_Value for new brand
        First_Open_Value = ws.Cells(i, 3).Value
        
    ElseIf ws.Cells(i, 1).Value = ws.Cells(j, 9).Value Then
        
        'Add to the total stock volume
        Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value
    
    End If

Next i


'Print info for the final brand

'Yearly change
ws.Cells(j, 10).Value = ws.Cells(Rows.Count, "F").End(xlUp) - First_Open_Value

'Percent change
ws.Cells(j, 11).Value = ws.Cells(j, 10).Value / First_Open_Value

'Total stock volume
ws.Cells(j, 12).Value = Total_Stock_Volume


'New loop to add conditional formatting to yearly change
For i = 2 To ws.Cells(Rows.Count, "A").End(xlUp).Row

    If ws.Cells(i, 10).Value > 0 Then
    
        ws.Cells(i, 10).Interior.ColorIndex = 4
    
    ElseIf ws.Cells(i, 10).Value < 0 Then
    
        ws.Cells(i, 10).Interior.ColorIndex = 3
    
    End If
    
Next i

'Challenge Work: 1

'Assume the first value is the greatest percent increase, the greatest percent decrease, and greatest total volume
ws.Cells(2, 16).Value = ws.Cells(2, 11).Value
ws.Cells(3, 16).Value = ws.Cells(2, 11).Value
ws.Cells(4, 16).Value = ws.Cells(2, 12).Value

'j will define the row number for column Ticker
For j = 2 To ws.Cells(Rows.Count, "I").End(xlUp).Row

'Check values against first value for percent change, replacing everytime one is higher than the new max or lower than the new min
    If ws.Cells(j, 11).Value > ws.Cells(2, 16).Value Then
        
        ws.Cells(2, 16).Value = ws.Cells(j, 11).Value
        ws.Cells(2, 15).Value = ws.Cells(j, 9).Value
    
    ElseIf ws.Cells(j, 11).Value < ws.Cells(3, 16).Value Then
    
        ws.Cells(3, 16).Value = ws.Cells(j, 11).Value
        ws.Cells(3, 15).Value = ws.Cells(j, 9).Value
    
    End If
    
Next j

'j will define the row number for column Ticker
For j = 2 To ws.Cells(Rows.Count, "I").End(xlUp).Row

'Check values against first value for total volume, replacing everytime one is higher than the new max
    If ws.Cells(j, 12).Value > ws.Cells(4, 16).Value Then
        
        ws.Cells(4, 16).Value = ws.Cells(j, 12).Value
        ws.Cells(4, 15).Value = ws.Cells(j, 9).Value
    
    End If
    
Next j

Next ws

End Sub
