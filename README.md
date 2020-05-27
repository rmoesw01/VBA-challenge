# VBA-challenge
VBA scripting to analyze real stock market data

##Data File Format requirements:
*	Macro-enabled excel format
*	All columns must have headers in the first row
*	Column Values:
	*	column A: ticker symbol 
	*	column B: date 
	*	column C: opening value
	*	column F: closing value
	*	column G: stock volume

*	columns H through Q should be blank

* 	Data must be sorted first by ticker column(A), then by date column(B) (smallest to largest)

**** CAUTION **** This program does not allow for skipped rows, please make sure your data is all together

##To run the script on your data file, complete the following steps:
1. open your macro-enabled excel file
2. sort all worksheets in your workbook to meet the above requirements
3. import VBAStocks_release.bas into your spreadsheet
4. select and run the Macro called VBAStocks2

##Output:
*	Columns H through K will output one row for each unique Ticker in the worksheet with the format shown below
	*	Column H: Ticker Symbol
	*	Column I: Yearly Change - the difference of the closing value at the end of the year and the opening value at the beginning of the year
	*	Column J: Percent Change - the percent change of the Yearly Change / opening value
	*	Column K: Total Stock Volume - the sum of the stock volume for the ticker for the year
*	Columns O through Q will output the greatest percent increase, percent decrease, and total stock volume for the worksheet with the format shown below
	*	Column O: labels
	*	Column P: Ticker Symbol
	*	Column Q: value of the greatest value