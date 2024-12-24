# DUBC_Module2: Quarterly Stock Analysis VBA Script
This is the repository for the 2nd challenge for the DU Data Analytics Bootcamp. 

This VBA script is designed to analyze quarterly stock data across multiple worksheets and generate summary tables. It calculates various stock statistics such as the percent change, total stock volume and price change for each ticker symbol. This script also identifies the stock ticker with the greatest percent increase, the greatest percent decrease and the greatest total volume.

Features
 - Opening Price: Calculates the opening price for each stock ticker for the quarter
 - Closing Price: Calculates the closing price for each stock ticker for the quarter
 - Price Change: Calculates the change in stock price over the quarter (closeing price - opening price)
 - Percent Change: Calculates the percentage change in stock price relative to the opengin price
 - Total Volume: Calculates the sum of the total trading volume for each stock
 - Max Percent Increase: Identifies the ticker with the greatest percent increase in stock price
 - Max Percent Decrease: Identifies the ticker with the greatest percent decrease in stock price
 - Max Total Volume: Identifies the ticker with the greatest total trading volume

Output
 - Summary Table: Displays the ticker symbol, price change, percent change and total volume for each ticker in a summary table starting at column I for each worksheet.
 - Max Statistics Table: Displays the stock with the greatest percent increase, greatest percent decrease and greatest trading volume in a new summary table starting at column O

 Prerequisites
 - Excel: This script is designed to run in Micorsoft Excel
 - Stock Data: the stock data worksheets 
    - Column A: Ticker symbol
    - Column B: Date
    - Column C: Opening Price
    - Column F: Closing Price
    - Column G: Volume Traded

Column Mapping
- Column A: Ticker symbol
- Column C: Opening Price
- Column F: Closing Price
- Column G: Volume Traded
- Column I: Ticker Symbol (for summary table)
- Column J: Quarterly Price Change
- Column K: Percent Change (rounded to 2 decimal places)
- Column L: Total Volume

Max Summary Table Columns
- Column o: Description ("Greatest % Increase", "Greatest % Decrease, "Greate Total Volume")
- Column P: Ticker Symbol corresponding to the max value
- Column Q: The max value (percent increase, percent decrease, total volume)

Instructions for Use
1. Open the workbook in Excel
2. Prepare Data: Make sure the worksheet(s) contain the stock data with the required columns as mentioned above (Ticker Symbol, Opening Price, Closing Price, and Volume). Make sure the ticker symbols are listed alphebetically (column A) and in chronological order (column B).
3. Insert VBA Code:
    - Open Visual Basic editor (from the Developer tab in Excel)
    - Go to Insert > Module to create a new module
    - Copy and paste this entire code into the new module
4. Run the script
5. Check the output

Troubleshooting
- Data Format: Ensure that the stock data is structured correctly in the required columns
- Empty Cells: If there are empty rows or invalid data, it may cause errors. Make sure all the required fields are filled in correctly.
