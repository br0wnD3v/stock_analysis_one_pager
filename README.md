# ONE PAGER STOCK INFORMATION

Allows you to create one-page information, usually important metrics about the input stock.

## REQURIREMENTS

1. Have OpenAI's API_KEY.
2. Have a `35stocks.xlsx` file - It should have the important metrics regarding the stock like P/E TTM, Dividend Yield etc etc. Example file is attached.
3. Have a `35stockscomparable.xlsx` file - List its comparable companies (Something closer to its industry or whatever.)

The use of the xlsx files : The important metrics for a stock as mentioned in `35stocks.xlsx`, they are important as in each company is different and can't be compared based on a fixed metric. The other file(`35stockscomparable.xlsx`) lets us get a mean value(of its peers) of the important metric selected and color code the selected important metric in the pdf file for the entered stock.
For all metrics other than Dividend Yield, the lesser the value from its mean the better (Hence green) and vica versa for Dividend Yield (better if more than mean hence green).

You can have as many important metric as you can, just make sure to add them under the existing metrics in the excel file or just use as you wish.

Note : The file name can be changed to whatever you wish, just make sure to change them in the .py file as well since its hard-coded.

### I WILL NOT BE MAINTAINIG THE CODE OR ANYTHING SO FEEL FREE TO TAKE OVER AS SEEN FIT.

# IMPORTANT

The project scrapes from yahoo finance and hence should not be be used for commercial purposes without their permission.
The project is ONLY intended for education purposes and we have no liability if used for anything else.

# ============ CODE IS LAW =============
