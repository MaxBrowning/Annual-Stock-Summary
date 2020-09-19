Attribute VB_Name = "Module1"
Sub VBA_Challenge()

'Define all variables
Dim First_Open_Value As Double
Dim Last_Close_Value As Double
Dim Yearly_Change As Double
Dim Total_Stock_Volume As Double

'Set the row two's values
Cells(2, 9).Value = Cells(2, 1).Value
First_Open_Value = Cells(2, 3).Value
Total_Stock_Volume = Cells(2, 7).Value

'j will define the row number for column Ticker
j = 2

'i will define the row number for column <ticker>
For i = 2 To Cells(Rows.Count, "A").End(xlUp).Row
    
    If Cells(i, 1).Value <> Cells(j, 9).Value Then
    
    'Go to the next row.
    j = j + 1
    
    'Print new ticker symbol
    Cells(j, 9).Value = Cells(i, 1).Value
    
    'Save last close value
    Last_Close_Value = Cells(i - 1, 6).Value
    
    'Calculate yearly change
    Cells(j - 1, 10).Value = Last_Close_Value - First_Open_Value
    Yearly_Change = Cells(j - 1, 10).Value
    
    'Print the previous brand's percent change
    Cells(j - 1, 11).Value = (Yearly_Change / First_Open_Value)
    
    'Print the previous brand's total stock volume
    Cells(j - 1, 12).Value = Total_Stock_Volume
    
    'Reset Total_Stock_Volume starting value for new brand
    Total_Stock_Volume = Cells(i, 7).Value
    
    'Reset First_Open_Value for new brand
    First_Open_Value = Cells(i, 3).Value
    
    ElseIf Cells(i, 1).Value = Cells(j, 9).Value Then
    
    'Add to the total stock volume
    Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7).Value
    
    End If

Next i


'Print info for the final brand

'Yearly change
Cells(j, 10).Value = Cells(Rows.Count, "F").End(xlUp) - First_Open_Value

'Percent change
Cells(j, 11).Value = Cells(j, 10).Value / First_Open_Value

'Total stock volume
Cells(j, 12).Value = Total_Stock_Volume


'New loop to add conditional formatting to yearly change
For i = 2 To Cells(Rows.Count, "A").End(xlUp).Row

    If Cells(i, 10).Value > 0 Then
    Cells(i, 10).Interior.ColorIndex = 4
    
    ElseIf Cells(i, 10).Value < 0 Then
    Cells(i, 10).Interior.ColorIndex = 3
    
    End If
    
Next i

End Sub
