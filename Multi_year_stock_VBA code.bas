Attribute VB_Name = "Module1"
Sub Stock_DataHW():

Dim Total_Vol As Double
Dim i As Long
Dim yearly_change As Single
Dim j As Integer
Dim initial_open As Long
Dim rowCount As Long
Dim percentChange As Single
Dim ws As Worksheet

For Each ws In Worksheets

j = 0
Total_Vol = 0
yearly_change = 0
initial_open = 2

'Set Title row
ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Volume"


'Get the row number of the last row with data
rowCount = Cells(Rows.Count, "A").End(xlUp).Row

For i = 2 To rowCount

If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    
    'Store results in variables
    Total_Vol = Total_Vol + ws.Cells(i, 7).Value
    
    'Zero total volume
    If Total_Vol = 0 Then
        'Print results
        ws.Range("I" & 2 + j).Value = ws.Cells(i, 1).Value
        ws.Range("J" & 2 + j).Value = 0
        ws.Range("K" & 2 + j).Value = "%" & 0
        ws.Range("L" & 2 + j).Value = 0
    Else
    
        'find first non-zero starting value
        If ws.Cells(initial_open, 3) = 0 Then
            For find_value = initial_open To i
                If ws.Cells(find_value, 3).Value <> 0 Then
                    initial_open = find_value
                    Exit For
                End If
            Next find_value
        End If
        
        'Calculate change
        yearly_change = (ws.Cells(i, 6) - ws.Cells(initial_open, 3))
        percentChange = Round((yearly_change / ws.Cells(initial_open, 3) * 100), 2)
        
        'Start of next ticker
        initial_open = i + 1
        
        'print results
        ws.Range("I" & 2 + j).Value = ws.Cells(i, 1).Value
        ws.Range("J" & 2 + j).Value = Round(yearly_change, 2)
        ws.Range("K" & 2 + j).Value = "%" & percentChange
        ws.Range("L" & 2 + j).Value = Total_Vol
        
        'Conditional formatting
        Select Case yearly_change
            Case Is > 0
                ws.Range("J" & 2 + j).Interior.ColorIndex = 4
            Case Is < 0
                ws.Range("J" & 2 + j).Interior.ColorIndex = 3
        End Select
    End If
    
    
    Total_Vol = 0
    yearly_change = 0
    j = j + 1
    
    'If ticker does not change add results
    Else
        Total_Vol = Total_Vol + ws.Cells(i, 7).Value
    End If
    
Next i

Next ws

    
End Sub

