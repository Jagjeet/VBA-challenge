# VBA-challenge

This project uses a VBA script to analyze stock market data. It goes through each sheet in an excel doc containing stock data and outputs the following information in the columns to the right of the data (See the sample results section below for more info):

* The ticker symbol.
* Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.
* The percent change from opening price at the beginning of a given year to the closing price at the end of that year.
* The total stock volume of the stock.

It also provides some conditional formatting that will highlight positive changes in green and negative changes in red

## Excel Document Format

Each sheet in the excel document is expected in the following format:

![image info](./images/format.png)

The columns below should exist in each sheet in columns A to G:

* ticker
* date
* open
* high
* low
* close
* volume

Note the exact titles of the columns are not required, but just that a header row is expected with the data expected below. Also the `ticker` data is expected in alphabetical order to save time processing.

## Usage

1. Copy and paste the contents of GenerateStockInfo.vba into `ThisWorkbook` in the VBAProjects for the excel document that you are processing.
2. Run the script by clicking the play button and running `ThisWorkbook.GenerateStockInformation`

## Sample Results

Sample results can be found in images/results for:

* [2014](./images/results/2014.jpg)
* [2015](./images/results/2015.jpg)
* [2016](./images/results/2016.jpg)

Note due to limitations in exporting images from excel only the first 512 lines appear in the screenshots. 

## References

The VBA script made use of concepts found at the following links:

### Best Practices/Commenting Styles

* https://www.experts-exchange.com/articles/21759/A-Guide-to-Writing-Understandable-and-Maintainable-VBA-Code.html

### Cell Coloring

* http://dmcritchie.mvps.org/excel/colors.htm

### Debugging

* https://www.automateexcel.com/vba/debug-print-immediate-window

### Looping through Excel Sheets

* https://excelhelphq.com/how-to-loop-through-worksheets-in-a-workbook-in-excel-vba/
* https://support.microsoft.com/en-us/topic/macro-to-loop-through-all-worksheets-in-a-workbook-feef14e3-97cf-00e2-538b-5da40186e2b0

### Subroutines/Functions

* https://docs.microsoft.com/en-us/office/vba/language/concepts/getting-started/calling-sub-and-function-procedures
* https://www.automateexcel.com/vba/function/
* https://www.excel-easy.com/vba/function-sub.html
* https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/function-statement
