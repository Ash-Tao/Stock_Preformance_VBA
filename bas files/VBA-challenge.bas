Attribute VB_Name = "Module1"
Sub VBAChallenge()

    Dim wb As Workbook
    Dim ws As Worksheet

    Set wb = ThisWorkbook
    Set ws = wb.ActiveSheet
    Set Rng = ws.Range("I1")
    Set rngsmpl = ws.Range("A1")
    
    'remove duplicated company name and copy to new column
    ws.Range("A:A").AdvancedFilter Action:=xlFilterCopy, CopyToRange:=Rng, Unique:=True
    
    'put header on, since i don't know how to run "xlFilterCopy" without header
    Cells(1, 9) = "Ticket"
    Cells(1, 10) = "Yearly Change"
    Cells(1, 11) = "Percent Chang"
    Cells(1, 12) = "Total Stock Volume"
    
    'add up total value based on the company name
    Dim EndRow As Long
    Dim EndRowSmpl As Long
    EndRow = Rng.End(xlDown).Row
    EndRowSmpl = rngsmpl.End(xlDown).Row
    
    For X = 2 To EndRow
    ws.Cells(X, 12) = WorksheetFunction.SumIf(ws.Range("A:A"), ws.Cells(X, 9), ws.Range("G:G"))
    Next X
    
    'convert date to number future comparison
    'Dim RngCvt As Range
    'For Each RngCvt In Range("B:B").Columns
    '    RngCvt.TextToColumns
    'Next RngCvt
       
    'find the price on earliest and latest date for each company
    Dim MinDate As Long
    Dim MaxDate As Long
    Dim EariestPrice As Double
    Dim LatestPrice As Double
    
    'find price
    For y = 2 To EndRow
        MinDate = Cells(2, 2).Value
        MaxDate = Cells(2, 2).Value
        'find eariest price
        For i = 2 To EndRowSmpl
            If Cells(i, 1).Value = Cells(y, 9).Value And Cells(i, 2).Value <= MinDate Then
                MinDate = Cells(i, 2).Value
                EariestPrice = Cells(i, 3).Value
            End If
        Next i
        'find latest price
        For h = 2 To EndRowSmpl
            If Cells(h, 1).Value = Cells(y, 9).Value And Cells(h, 2).Value >= MaxDate Then
                MaxDate = Cells(h, 2).Value
                LatestPrice = Cells(h, 6).Value
            End If
        Next h
        
        'format cells
        Cells(y, 10) = Format(LatestPrice - EariestPrice, "#,##0.00")
        Cells(y, 11) = Format(Cells(y, 10) / EariestPrice, "#,##0.00%")

    Next y
    
        'conditional formatting
        'remove existing conditional formatting from the range
        Range("J:J").FormatConditions.Delete
        'negative in red
        Range("J:J").FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, _
        Formula1:="=0"
        Range("J:J").FormatConditions(1).Interior.Color = vbRed
        'negative in green
        Range("J:J").FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, _
        Formula1:="=0"
        Range("J:J").FormatConditions(2).Interior.Color = vbGreen
        'remove conditional formatting from header
        Range("J1:J1").ClearFormats

End Sub





