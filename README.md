# VBA-Assignment

# Background

In this activity, I utilized VBA to manipulate and elucidate data from a spreadsheet of real-world stock data from 2014-2016. The workbook which was used was titled Multiple_year_stock_data.xlsx, and its size prevents it from being accessed directly on GitHub. Test data (alphebetical_testing.xlsx) was also used to develop the scripts.

Microsoft Office Home and Student 2019 provided the Excel application which was utilized in Visual Basic Coding. Additionally, I used a PC with Windows 10.

# What did I do with these data?

The code is held in the file VBA_Stock_Calculations.bas within this repository. Once this is downloaded, the steps to run this program after opening up the multi-year stock data workbook are as follows:
- Open Visual Basic in Excel through the Developer tab
- Import the .bas file using **File -> Import File**
- Run the macro by pressing the green play button on the top menu of the interface

After following these steps, the script will loop through each worksheet in the file and output various metrics for characterizing the stocks.

Here is what the script outputs on each page of the workbook:
- The ticker associated with each stock
- Yearly change from opening price to closing price
- Percent change from the beginning of the year to end of the year (closing price - opening price)
- Total volume of stock
- Conditional formatting on the yearly changes with green, yellow, and red to indicate growth, stagnation, or decline in stock price
- A separate return of the stock with the greatest percent increase, greatest percent decrease, and greatest total volume

#Script Output

The results of the script can be viewed in the **Screenshots** folder of this repository. 




