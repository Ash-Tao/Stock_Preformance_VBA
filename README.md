# **VBA-Challenge: The VBA for Stock Market Analyst**

## Purpose
Use `VBA` scripting to analyse generated stock market data.

### **Target (Module 1)**
Create a script that loops through all the stocks for one year and outputs the following information:<br />
- Get the ticker symbol.<br />
  Use `AdvancedFilter` to remove duplicated values and copy to a specific column (column I). <br />
    ```diff
    Dim wb As Workbook
    Dim ws As Worksheet
    Set Rng = ws.Range("I1")
    ws.Range("A:A").AdvancedFilter Action:=xlFilterCopy, CopyToRange:=Rng, Unique:=True
    
- Find Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.<br />
   In the real world, it might be possible to come across listing or delisting in the middle of a year. Or the data is not presented in chronological order   after the company name as a group. Need to find a way to get accurate data in a disordered range.<br />
   Use nested loop to find the difference between closing and beginning price for each stock<br />
A variable range is applied. `Range.End (xlDown)`<br />
   Dim four variables for the found closing and beginning price. <br />
   Do the calculation. <br />
    ```diff
    Dim MinDate As Long
    Dim MaxDate As Long
    Dim EariestPrice As Double
    Dim LatestPrice As Double
     
    Set Rng = ws.Range("I1")
    Set rngsmpl = ws.Range("A1")
    Dim EndRow As Long
        EndRow = Rng.End(xlDown).Row
    Dim EndRowSmpl As Long
        EndRowSmpl = rngsmpl.End(xlDown).Row

    For y = 2 To EndRow
        MinDate = Cells(2, 2).Value
        MaxDate = Cells(2, 2).Value
        
    #find eariest price.
        For i = 2 To EndRowSmpl
            If Cells(i, 1).Value = Cells(y, 9).Value And Cells(i, 2).Value <= MinDate Then
                MinDate = Cells(i, 2).Value
                EariestPrice = Cells(i, 3).Value
            End If
        Next i
        
    #find latest price.
        For h = 2 To EndRowSmpl
            If Cells(h, 1).Value = Cells(y, 9).Value And Cells(h, 2).Value >= MaxDate Then
                MaxDate = Cells(h, 2).Value
                LatestPrice = Cells(h, 6).Value
            End If
        Next h
        
    #calculate and format cells
        Cells(y, 10) = Format(LatestPrice - EariestPrice, "#,##0.00")

- The percent change from opening price at the beginning of a given year to the closing price at the end of that year.<br />
  Use the results above, do the calculation and format the cells
    ``` diff
        Cells(y, 11) = Format(Cells(y, 10) / EariestPrice, "#,##0.00%")

- The total stock volume of the stock.<br />
  Use `WorksheetFunction.SumIf` to get the total volume for each stock.
    ``` diff
        For X = 2 To EndRow
            ws.Cells(X, 12) = WorksheetFunction.SumIf(ws.Range("A:A"), ws.Cells(X, 9), ws.Range("G:G"))
        Next X

- Highlights positive change in green and negative change in red<br />
  use `Conditional Formatting`
    ``` diff
        Range("J:J").FormatConditions.Delete 
    #negative in red
        Range("J:J").FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, _
        Formula1:="=0"
        Range("J:J").FormatConditions(1).Interior.Color = vbRed
    #negative in green
        Range("J:J").FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, _
        Formula1:="=0"
        Range("J:J").FormatConditions(2).Interior.Color = vbGreen
    #remove conditional formatting from header
        Range("J1:J1").ClearFormats

### **Bonus (Module 2)**
- Return the stock with the "Greatest % increase", "Greatest % decrease" and "Greatest total volume" based on the data found.<br />
  - use `WorksheetFunction.Max` or `WorksheetFunction.Min` find the value.<br />
    ``` diff
    #find max & min value and format them
        ws.Cells(2, 16) = Format(WorksheetFunction.Max(ws.Range("K:K")), "#,##0.00%")
        ws.Cells(3, 16) = Format(WorksheetFunction.Min(ws.Range("K:K")), "#,##0.00%")
        ws.Cells(4, 16) = WorksheetFunction.Max(ws.Range("L:L"))

  - use loop to find the ticker.
    ``` diff
    #max on %
        For y = 2 To EndRow
            If Cells(y, 11).Value = Cells(2, 16).Value Then
               Cells(2, 15).Value = Cells(y, 9).Value
            End If
        Next y
    #min on %
        For y = 2 To EndRow
            If Cells(y, 11).Value = Cells(3, 16).Value Then
               Cells(3, 15).Value = Cells(y, 9).Value
            End If
        Next y
    #max on volume
        For y = 2 To EndRow
            If Cells(y, 12).Value = Cells(4, 16).Value Then
               Cells(4, 15).Value = Format(Cells(y, 9).Value, "##0.0E+0")
            End If
        Next y

## **Template screenshot**
The result should match the following image:<br />
![alt text](https://github.com/Ash-Tao/VBA-challenge/blob/main/Screen%20Shot/2018%201-3.png)
- [More screenshots](https://github.com/Ash-Tao/VBA-challenge/tree/main/Screen%20Shot)
- [The full results across 3 years](https://github.com/Ash-Tao/VBA-challenge/tree/main/Full%20Results)



## **How to Run**
- `.bas` file. <br />
  Download the `.bas` file.<br />
  [.bas files](https://github.com/Ash-Tao/VBA-challenge/tree/main/bas%20files)<br />
- Final report on full sample datasets.<br />
  Download `VBA Challenge_MultipleYearStock_data.xlsm` to your local drive.<br />
  [VBA Challenge_MultipleYearStock_data.xlsm](https://github.com/Ash-Tao/VBA-challenge/blob/main/VBA%20Challenge_MultipleYearStock_data.xlsm)
- Macros Button.<br />
  In each sheet, there are `VBA-Challenge` & `Bonus` two buttons. 
  They have been linked to the modules, which acts the same on every sheet.<br />
  ![image](https://github.com/Ash-Tao/VBA-challenge/blob/main/Macros%20Button.png)
  ```diff
  -NOTES: Please delete existing calculation result(from column I to column P) before press the Macros Button.
  -Otherwise, an error message of "Run-time error '1004' will occur and stop the code to be run.
- Small sample dataset - `alphabetical_testing.xlsx`.<br />
  If the full sample dataset is too large to loading. A sample of this small data can be downloaded for testing purposes.<br />
  [alphabetical_testing.xlsx](https://github.com/Ash-Tao/VBA-challenge/blob/main/alphabetical_testing.xlsm)

