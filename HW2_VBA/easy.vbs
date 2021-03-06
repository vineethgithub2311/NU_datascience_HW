Attribute VB_Name = "Module1"
Sub Stocks()
    
' Dimensions
Dim total As Double
Dim j As Integer
'using worksheets for navigating within the file
Dim ws As Worksheet

    For Each ws In Worksheets
    ' Variables for each sheet (looping through all years)
    total = 0
    j = 0

    ' Title row setup
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Total Stock Volume"

    ' To have the code identify the last row
    rCount = Cells(Rows.Count, "A").End(xlUp).Row
        
        For i = 2 To rCount
        ' Once new ticker is found, will print out results in new row
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                total = total + ws.Cells(i, 7).Value
                ws.Range("I" & 2 + j).Value = ws.Cells(i, 1).Value
                ws.Range("J" & 2 + j).Value = total
                total = 0
                j = j + 1
            Else
                total = total + ws.Cells(i, 7).Value

            End If

        Next i
        ' variable in worksheet are set to 0 before proceeding to next sheet
        total = 0
        j = 0

    Next ws


End Sub

