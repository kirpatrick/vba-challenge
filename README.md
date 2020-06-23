# vba-challenge
Analysis of stock market data using VBA

<!-- <img src="images/under-construction.png" alt="drawing" width="500"/> -->

# Project Overview
## SMU Data Science Boot Camp - Visual Basic Activity

### Instructions / Functionality / Design

#### Given
Stock Market data, in Microsoft Excel format, for multiple years.  Each tab represents a different year of data and contains daily company stock information as follows:
- Ticker symbol
- Date
- Opening Price
- Highest value
- Lowest value
- Closing value
- Trade Volume

#### Instructions and Design
* Create a VB script that will loop through all the stocks for one year for each run and take the following information.
    - The ticker symbol.
    - Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.
    - The percent change from opening price at the beginning of a given year to the closing price at the end of that year.
    - The total stock volume of the stock.
* You should also have conditional formatting that will highlight positive change in green and negative change in red.

#### Challenges
1.	Your solution will also be able to return the stock with the "Greatest % Increase", "Greatest % Decrease" and "Greatest Total volume".
2.	Make the appropriate adjustments to your VBA script that will allow it to run on every worksheet, i.e., every year, just by running the VBA script once.

### Source
[SMU DS Boot Camp - VBA Scripting Challenge](https://smu.bootcampcontent.com/SMU-Coding-Bootcamp/SMU-DAL-DATA-PT-11-2019-U-C/tree/master/02-Homework/02-VBA-Scripting/Instructions)

# Solution
1. Download the Excel workbook [vba-challenge/VBAStocks/Multiple_year_stock_data_analysis.xlsm](https://github.com/kirpatrick/vba-challenge/blob/master/VBAStocks/Multiple_year_stock_data_analysis.xlsm).
2. Open the downloaded workbook.  This may take a few seconds given the file's size.
3. Enable Macros.
4. See the Instructions' tab in the workbook for VB script execution.

[Multi Year Stock Data Screenshots](https://github.com/kirpatrick/vba-challenge/blob/master/VBAStocks/MultiYear_Stock_Analysis%E2%80%93Summary_Screenshots.pdf)

Individual execution scripts can be found in the [vba-challenge/VBAStocks](https://github.com/kirpatrick/vba-challenge/tree/master/VBAStocks) directory.

# Tech Stack
- Microsoft Excel 2016
