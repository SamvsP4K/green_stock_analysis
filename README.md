# Green Stock Analysis

## Purpose:
The purpose of this analysis was to create a VBA Macro about which stocks would be best for Steve's parents to invest in based on information from 2017 and 2018. 

## Data & Process:

The data we recieved was an excel document that contained rows of stock stickers, dates, volumes and open/close prices, and stock highs and lows.
Throughout the module we wrote VBA code that would iterate through all the stock tickers of one year and color code each ticker based on performance. 

The code was refactored to iterate through the tickers based on a year chosen by the user. For this workbook the only options were 2017 and 2018.

## Benefits:

A couple of the benfits of the refactored code is that it runs quicker than the original code and it can be used for multiple sets of data. If a new workbook was added to this the code could be easily applied to new the new data.

Below are examples snippets of the code and how fast each runs:

The original code run time:

![original_codetime](https://user-images.githubusercontent.com/110923091/203635575-848e967f-9654-448f-8d7a-b3d71d6784f7.PNG)

The refactored code run time:

![refactored_speed](https://user-images.githubusercontent.com/110923091/203635731-0d255cea-76ac-482e-9d6c-e03ae1bab3f3.PNG)




