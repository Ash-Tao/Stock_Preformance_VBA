# **VBA-Challenge: The VBA for Stock Market Analyst**

## Purpose
Use `VBA` scripting to analyze generated stock market data.

### **Target**
Create a script that loops through all the stocks for one year and outputs the following information:
- Get the ticker symbol.<br />
Use `AdvancedFilter` to remove duplicated values. <br /> 

- Find Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.<br />
Use nested loop to find the difference between closing and beginning price for each stock<br />
A variable range is applied. `Range.End (xlDown)`<br />

- The percent change from opening price at the beginning of a given year to the closing price at the end of that year.<br />
Dim four variables for the found closing and beginning price.<br />
Do the calculation.<br />
  - MinDate
  - EariestPrice
  - MaxDate
  - LatestPrice<br />

*In the real world, it might be possible to come across listings or delistings in the middle of a year. Or the data is not presented in chronological order after company name as a group. This script of mine can still find the data of each company in the disordered record*<br />

- The total stock volume of the stock.<br />
Usd `WorksheetFunction.SumIf` to get the total volume for each stock.

- Highlights positive change in green and negative change in red<br />
use `Conditional Formatting`


### **Bonus**
Return the stock with the "Greatest % increase", "Greatest % decrease" and "Greatest total volume" based on the data found.<br />
use `WorksheetFunction.Max` or `WorksheetFunction.Min` find the value<br />
use loop to find the ticker


## **Template screenshot**
The result should match the following image:
![alt text](https://github.com/Ash-Tao/VBA-challenge/blob/main/VBA%20Challenge%20Screen%20Shot/Screen%20Shot-Year%202018%201:3.png)

## **Files**
- [VBA Scripts On With Datasets](https://github.com/Ash-Tao/VBA-challenge/blob/main/2%20VBA%20Challenge_MultipleYearStock_data.xlsm)
- [VBA Challenge Screen Shot](https://github.com/Ash-Tao/VBA-challenge/tree/main/VBA%20Challenge%20Screen%20Shot)
