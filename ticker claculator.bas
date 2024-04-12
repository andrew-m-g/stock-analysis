Attribute VB_Name = "Module1"
Sub tickertracker()
Dim ws As Worksheet
'setting up data variables


'formatting each page with headers
For Each ws In Worksheets
ws.Range("I1") = "Ticker"
ws.Range("J1") = "Yearly Change"
ws.Range("K1") = "Percent Change"
ws.Range("L1") = "Total Stock Volume"

ws.Range("P1") = "Ticker"
ws.Range("Q1") = "Value"
ws.Range("O2") = "Greatest % Increase"
ws.Range("O3") = "Greatest % Decrease"
ws.Range("O4") = "Greatest Total Volume"
ws.Range("Q2:Q3").NumberFormat = "0.00%"

Dim ticker As String
Dim tsv As Double
Dim openV As Double
Dim closingV As Double

rc = ws.Cells(Rows.Count, 1).End(xlUp).Row
For x = 2 To rc 'collect data for ticker
            If ws.Cells(x - 1, 1).Value <> ws.Cells(x, 1).Value Then
            ticker = ws.Cells(x, 1).Value
            openV = ws.Cells(x, 3).Value
            tsv = tsv + ws.Cells(x, 7).Value

        'collecting all values for stock volume
            ElseIf ws.Cells(x - 1, 1).Value = ws.Cells(x, 1).Value Then
            tsv = tsv + ws.Cells(x, 7).Value
        'checking final value within ticker set + finding yearly change
            If ws.Cells(x + 1, 1).Value <> ws.Cells(x, 1).Value Then
            closingV = ws.Cells(x, 6).Value
            lastV = ws.Cells(Rows.Count, 10).End(xlUp).Row + 1
            ws.Cells(lastV, 9).Value = ticker
            ws.Cells(lastV, 10).Value = openV - closeV
                 If ws.Cells(lastV, 10).Value >= 0 Then
                 ws.Cells(lastV, 10).Interior.ColorIndex = 4
                 ElseIf ws.Cells(lastV, 10).Value < 0 Then
                 ws.Cells(lastV, 10).Interior.ColorIndex = 3
                 End If
        'calculate percentage change
            ws.Cells(lastV, 11).Value = (openV - closeV) / openV
            ws.Cells(lastV, 11).NumberFormat = "0.00%"
            End If
            End If
Next x

'find greatest % increase + decrease and largest stock volume
Next
End Sub
