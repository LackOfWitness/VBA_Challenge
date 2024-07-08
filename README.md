The VBA script `TickerPriceMacro` performs a comprehensive analysis of stock prices across multiple worksheets in an Excel workbook. Below is a breakdown of the steps involved:

# Stock Data Processing Macro

## Step 1: Create the CONCAT Column
The macro creates a new column H and populates it with concatenated values of ticker (column A) and date (column B) in the format "mm/dd/yyyy".

## Step 2: Copy Unique Tickers to Column J
The macro uses a dictionary to collect unique tickers from column A and writes them to column J.

## Step 3: Calculate Earliest and Latest Dates
For each unique ticker, the macro calculates the earliest and latest dates, corresponding opening and closing prices, and total stock volume:

- It iterates through the rows to find the minimum and maximum dates for each ticker.
- It also calculates the total volume for each ticker.
- The results are written to columns K (open date), L (close date), M (open price), N (close price), O (quarterly change), P (percent change), and Q (total stock volume).

## Step 4: Apply Conditional Formatting
The macro applies conditional formatting to columns O "Quarterly Change"

- Green color for positive values.
- Red color for negative values.

## Step 5: Summary Statistics
The macro calculates summary statistics for the greatest percentage increase, decrease, and total volume:

- It iterates through the tickers to find the maximum and minimum percentage changes and the maximum total volume.
- The results are written to columns S (Greatest % Increase), T (Greatest % Decrease), and U (Greatest Total Volume).

At the end of the script, a message box notifies the user that the "Ticker Price Macro has completed successfully for all worksheets!", indicating the completion of the analysis.

## Q1 Screen Shot Results

![alt text](<Screenshot Q1.png>)

## Q2 Screen Shot Results

![alt text](<Screenshot Q2.png>)

## Q3 Screen Shot Results 

![alt text](<Screenshot Q3.png>)

## Q4 Screen Shot Results

![alt text](<Screenshot Q4.png>)

