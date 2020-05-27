# VBA-challenge
VBA scripting to analyze real stock market data

Data File Format requirements:
<br><TAB INDENT=5>	Macro-enabled excel format
<br><TAB INDENT=5>	All columns must have headers in the first row
<br><TAB INDENT=5>	Column Values:
<br><TAB INDENT=10>		column A: ticker symbol 
<br><TAB INDENT=10>		column B: date 
<br><TAB INDENT=10>		column C: opening value
<br><TAB INDENT=10>		column F: closing value
<br><TAB INDENT=10>		column G: stock volume
<br>
<br><TAB INDENT=5>	columns H through Q should be blank
<br>
<br><TAB INDENT=5> 	Data must be sorted first by ticker column(A), then by date column(B) (smallest to largest)
<br>
<br>**** CAUTION **** This program does not allow for skipped rows, please make sure your data is all together
<br>
<br>To run the script on your data file, complete the following steps:
<br>1. open your macro-enabled excel file
<br>2. sort all worksheets in your workbook to meet the above requirements
<br>3. import VBAStocks_release.bas into your spreadsheet
<br>4. select and run the Macro called VBAStocks2
<br>
<br>Output:
<br><TAB INDENT=5>	Columns H through K will output one row for each unique Ticker in the worksheet with the format shown below
<br><TAB INDENT=10>		Column H: Ticker Symbol
<br><TAB INDENT=10>		Column I: Yearly Change - the difference of the closing value at the end of the year and the opening value at the beginning of the year
<br><TAB INDENT=10>		Column J: Percent Change - the percent change of the Yearly Change / opening value
<br><TAB INDENT=10>		Column K: Total Stock Volume - the sum of the stock volume for the ticker for the year
<br><TAB INDENT=5>	Columns O through Q will output the greatest percent increase, percent decrease, and total stock volume for the worksheet with the format shown below
<br><TAB INDENT=10>		Column O: labels
<br><TAB INDENT=10>		Column P: Ticker Symbol
<br><TAB INDENT=10>		Column Q: value of the greatest value