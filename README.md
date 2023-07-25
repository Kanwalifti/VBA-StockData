# VBA-Challenge

#Background
This project involves using VBA scripting to analyze real stock market data. The analysis is performed using two sets of data: Test Data and Stock Data. The Test Data, used during script development, consists of 7 sheets (A to P) and is relatively smaller in size, serving as a testing dataset. On the other hand, the Stock Data, the main data used to run the script, comprises three sheets categorized by year (2018, 2019, and 2020). This file is larger in size and may require more time to execute the script.

Both sets of data are sourced into Microsoft Excel. The VBA scripts are available in the directories of both data points' files.

# Solution:

The VBA script provides the following functionalities when analyzing the stock market data:

# Solution 1: Ticker Symbol
The script sorts the distinct ticker symbols in column "I" with a column header "Ticker."

# Solution 2: Yearly Change
The script calculates the yearly change, which is the difference between the opening price at the beginning of a given year and the closing price at the end of that year. The values are displayed in column "J."

# Solution 3: Percent Change
The script calculates the percent change, which represents the percentage difference between the opening price at the beginning of a given year and the closing price at the end of that year. The values are displayed in column "K."

# Solution 4: Total Stock Volume
The script generates the total stock volume and presents it in column "L."

# Solution 5: Greatest Performances
The script also identifies and displays the "Greatest % Increase," "Greatest % Decrease," and "Greatest Total Volume" for the stocks analyzed. These values are provided in the analysis.

The script efficiently loops through all the stocks data, providing the necessary information and performing the specified calculations. Additionally, it includes conditional formatting to highlight positive changes in green and negative changes in red for better data visualization.
