# **VBA-challenge: The VBA for Stock Market Analyst**

## Purpose
Use `VBA` scripting to analyze generated stock market data.

### **Target**
Create a script that loops through all the stocks for one year and outputs the following information:
- Get the ticker symbol.<br />
Use `AdvancedFilter` to remove duplicated value. <br /> 

- Find Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.<br />
Use nested loop to find the difference between closing and beginning price for each stock<br />
A variable range is applied. `Range.End (xlDown)`<br />

- Fhe percent change from opening price at the beginning of a given year to the closing price at the end of that year.<br />
Dim two variables for the found closing and beginning price.<br />
Do the calculation.<br />

- The total stock volume of the stock.<br />
Usd `WorksheetFunction.SumIf` to get the total volume for each stock.

- Highlights positive change in green and negative change in red<br />
use `Conditional Formatting`


### **Bonus**
Return the stock with the "Greatest % increase", "Greatest % decrease" and "Greatest total volume" based on the data found.<br />
use `WorksheetFunction.Max` find the value<br />
use loop to find the ticker


## **Template screenshot**
The result should match the following image:


## **Files**
- VBA scripts on with datasets.
- Screenshot for partially result Y2018
- Screenshot for partially result Y2019
- Screenshot for partially result Y2020

