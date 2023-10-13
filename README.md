# VBA-challenge

## Task

Use VBA scripting to analyze generated stock market data. 

Create a script that loops through all the stocks for one year and outputs the following information:

* The ticker symbol

* Yearly change from the opening price at the beginning of a given year to the closing price at the end of that year.

* The percentage change from the opening price at the beginning of a given year to the closing price at the end of that year.

* The total stock volume of the stock.

* Add functionality to your script to return the stock with the "Greatest % increase", "Greatest % decrease", and "Greatest total volume". The solution should match the following image:

* Make the appropriate adjustments to your VBA script to enable it to run on every worksheet (that is, every year) at once.

## Description

I started by declaring the variables. I then got the names of the worksheets. Found the last row and skipped the header row and reset the ticker's count. I printed the headers. I sorted the data, started the loop, declared the starting ticker based on the current stock and stored the next ticker. Then I calculated the yearly and percentage changes. Accounted for the stock volume and moved on to the next value. I formatted the other columns. Printed the column headers for the analysis. Sorted the greatest percentage. Printed the greatest increase and decrease percentage. Sorted to find the volume and printed the greatest volume. Formatted the remaining columns. Added functionality to switch to the next worksheet.

## References

Reference for sorting with VBA

https://trumpexcel.com/sort-data-vba/

Reference for formatting in VBA

https://www.automateexcel.com/vba/format-cells/
