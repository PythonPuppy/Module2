Attribute VB_Name = "Module1"
Sub StockAnalysis()

' Loop Through the worksheets in Workbook
For Each ws In ThisWorkbook.Worksheets


'Variable Declaration

Dim Ticker As String
Dim Yearly_Change As Double
Dim Percent_Change As String
Dim total_Volume As Double
Dim opening_price As Double
Dim closing_price As Double
Dim Max_Percent As Double
Dim Min_Percent As Double
Dim Ticker1 As String
Dim Ticker2 As String
Dim Ticker3 As String
Dim Greatest_Volume As Double



'To store the value of last row in each sheet
lastRow = ws.Cells(Rows.Count, "A").End(xlUp).Row

'Initialize Total Volume
total_Volume = 0


'Initialize Summary_table
Dim summary_table_row As Integer
summary_table_row = 2

'''''''''Code to find the Yearly change , percent change based on Year and ticker'''''''''''''''''''''''''''''''''''''''''''''''''''

For i = 2 To lastRow
 
If ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then
    opening_price = ws.Cells(i, 3).Value
    
End If

    total_Volume = total_Volume + ws.Cells(i, 7).Value

If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

     Ticker = ws.Cells(i, 1).Value
     closing_price = ws.Cells(i, 6).Value
    
     Yearly_Change = closing_price - opening_price

If opening_price <> 0 Then

     Percent_Change = Format(((Yearly_Change / opening_price)), "0.00%")
     
Else
 
     Percent_Change = 0

End If

total_Volume = total_Volume + ws.Cells(i, 7).Value
ws.Range("I" & summary_table_row).Value = Ticker


ws.Range("J" & summary_table_row).Value = Yearly_Change
If ws.Range("J" & summary_table_row).Value < 0 Then
    ws.Range("J" & summary_table_row).Interior.ColorIndex = 3
Else
     ws.Range("J" & summary_table_row).Interior.ColorIndex = 4
End If

ws.Range("K" & summary_table_row).Value = Percent_Change
If ws.Range("K" & summary_table_row).Value < 0 Then
    ws.Range("K" & summary_table_row).Interior.ColorIndex = 3
Else
    ws.Range("K" & summary_table_row).Interior.ColorIndex = 4
    
End If
ws.Range("L" & summary_table_row).Value = total_Volume
summary_table_row = summary_table_row + 1
total_Volume = 0



End If
Next i







'---------------------Loop through the Summary table to find the greatest , Lowest percent change and volume --------------------


'Loop for Max_Percent

'Initialize variables
Max_Percent = 0


For a = 2 To 3001

If ws.Cells(a, 11).Value > Max_Percent Then

        Max_Percent = ws.Cells(a, 11)
        ws.Cells(2, 17).Value = Max_Percent
        Ticker1 = ws.Cells(a, 9)
        ws.Cells(2, 16).Value = Ticker1



End If
Next a



'Loop for Min_Percent

'Initialize variables
Min_Percent = 0


For b = 2 To 3001

If ws.Cells(b, 11).Value < Min_Percent Then

        Min_Percent = ws.Cells(b, 11)
        ws.Cells(3, 17).Value = Min_Percent
        Ticker2 = ws.Cells(b, 9)
        ws.Cells(3, 16).Value = Ticker2



End If
Next b



'Loop for greatest Total Volume

'Initialize variables
Greatest_Volume = 0


For c = 2 To 3001

If ws.Cells(c, 12).Value > Greatest_Volume Then

        Greatest_Volume = ws.Cells(c, 12)
        ws.Cells(4, 17).Value = Greatest_Volume
        Ticker3 = ws.Cells(c, 9)
        ws.Cells(4, 16).Value = Ticker3



End If
Next c


Next ws



End Sub
