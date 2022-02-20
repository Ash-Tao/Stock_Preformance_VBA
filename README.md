# **VBA-Challenge: The VBA for Stock Market Analyst**

## Purpose
Use `VBA` scripting to analyse generated stock market data.

### **Target (Module 1)**
Create a script that loops through all the stocks for one year and outputs the following information:<br />
- Get the ticker symbol.<br />
Use `AdvancedFilter` to remove duplicated values. <br />

- Find Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.<br />
Use nested loop to find the difference between closing and beginning price for each stock<br />
A variable range is applied. `Range.End (xlDown)`<br />

- The percent change from opening price at the beginning of a given year to the closing price at the end of that year.<br />
  In the real world, it might be possible to come across listing or delisting in the middle of a year. Or the data is not presented in chronological order   after the company name as a group. Need to find a way to get accurate data in a disordered range.<br />
  Dim four variables for the found closing and beginning price. <br />
  Do the calculation. <br />
  - MinDate
  - EariestPrice
  - MaxDate
  - LatestPrice<br />

- The total stock volume of the stock.<br />
Use `WorksheetFunction.SumIf` to get the total volume for each stock.

- Highlights positive change in green and negative change in red<br />
use `Conditional Formatting`


### **Bonus (Module 2)**
- Return the stock with the "Greatest % increase", "Greatest % decrease" and "Greatest total volume" based on the data found.<br />
  - use `WorksheetFunction.Max` or `WorksheetFunction.Min` find the value.<br />
  - use loop to find the ticker.


## **Template screenshot**
The result should match the following image:
![alt text](https://github.com/Ash-Tao/VBA-challenge/blob/main/VBA%20Challenge%20Screen%20Shot/Screen%20Shot-Year%202018%201:3.png)
More screenshots you can find here:
[VBA Challenge Screen Shot](https://github.com/Ash-Tao/VBA-challenge/tree/main/VBA%20Challenge%20Screen%20Shot)

## **How to Run**
- `.bas` file. <br />
  Download the `.bas` file.<br />
  [.bas files](https://github.com/Ash-Tao/VBA-challenge/tree/main/bas%20files)<br />
- Final report on full sample datasets.<br />
  Download `VBA Challenge_MultipleYearStock_data.xlsm` to your local drive.<br />
  [VBA Challenge_MultipleYearStock_data.xlsm](https://github.com/Ash-Tao/VBA-challenge/blob/main/2%20VBA%20Challenge_MultipleYearStock_data.xlsm)
- Macros Button.<br />
  In each sheet, there are `VBA-Challenge` & `Bonus` two buttons. 
  They have been linked to the modules, which acts the same on every sheet.<br />
  ![image](https://github.com/Ash-Tao/VBA-challenge/blob/main/Macros%20Button.png)
  ```diff
  -NOTES: Please delete existing calculation result(from column I to column P) before press the Macros Button.
  -Otherwise, an error message of "Run-time error '1004' will occur and stop the code to be run.
- Small sample dataset - `alphabetical_testing.xlsx`.<br />
  If the full sample dataset is too large to loading. A sample of this small data can be downloaded for testing purposes.<br />
  [alphabetical_testing.xlsx](https://github.com/Ash-Tao/VBA-challenge/upload)

