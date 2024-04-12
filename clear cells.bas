Attribute VB_Name = "Module2"
Sub tickertracker_delete_button()
'erase prior attempts
Dim ws As Worksheet
For Each ws In Worksheets
ws.Range("I:W") = ""
Next
End Sub
