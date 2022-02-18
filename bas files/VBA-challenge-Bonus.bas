Attribute VB_Name = "Module2"
Sub Bonus()

    Dim wb As Workbook
    Dim ws As Worksheet
    Set wb = ThisWorkbook
    Set ws = wb.ActiveSheet
    Set Rng = ws.Range("I1")
    Dim EndRow As Long
    EndRow = Rng.End(xlDown).Row
    
    Dim MaxP As Double
    Dim MinP As Double
    Dim MaxV As Double
    
    'put header on
    Cells(2, 14) = "Greatest % Increase"
    Cells(3, 14) = "Greatest % Drease"
    Cells(4, 14) = "Greatest Total Valume"
    Cells(1, 15) = "Ticker"
    Cells(1, 16) = "Value"
    
    'find max & min value and format them
    ws.Cells(2, 16) = Format(WorksheetFunction.Max(ws.Range("K:K")), "#,##0.00%")
    ws.Cells(3, 16) = Format(WorksheetFunction.Min(ws.Range("K:K")), "#,##0.00%")
    ws.Cells(4, 16) = WorksheetFunction.Max(ws.Range("L:L"))
    
    'find company name based on value
    'max on %
    For y = 2 To EndRow
        If Cells(y, 11).Value = Cells(2, 16).Value Then
            Cells(2, 15).Value = Cells(y, 9).Value
        End If
    Next y
    'min on %
    For y = 2 To EndRow
        If Cells(y, 11).Value = Cells(3, 16).Value Then
            Cells(3, 15).Value = Cells(y, 9).Value
        End If
    Next y
    'max on volume
    For y = 2 To EndRow
        If Cells(y, 12).Value = Cells(4, 16).Value Then
            Cells(4, 15).Value = Format(Cells(y, 9).Value, "##0.0E+0")
        End If
    Next y

End Sub


