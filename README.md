# vba-challenge
Analysis of stock market data using VBA

<!-- <img src="images/under-construction.png" alt="drawing" width="500"/> -->

# Project Overview
## SMU Data Science Boot Camp - Visual Basic Activity

### Instructions / Functionality / Design

#### Given
Stock data, in Microsoft Excel format, for multiple years.  Each tab represents a different year of data and contains daily company stock information as follows:
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
1. Download the Excel workbook *'Multiple_year_stock_data_analysis.xlsm'* from this repository.
2. Open the workbook.  This make take a few seconds given the file's size.
3. Enable Macros.
4. Follow the directions on the 'Instructions' tab.

# Tech Stack
- Microsoft Excel 2016 or later
