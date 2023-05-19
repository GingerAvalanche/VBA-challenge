# Homework Assignment 2 - Stock Analysis VBA Script

## Description

This script, as per the instructions, loops through each sheet in a given worksheet, reading ticker symbols, open and closing prices, and trade volumes, in order to provide summaries for each ticker, highlighting growth in green and shrink in red. It then finds the greatest % increases and decreases, and the greatest total volume, and outputs those in a separate secton of the sheet.

I wasn't aware of just how much the script was supposed to do, versus what I was supposed to do, outside of the script. As such, the script handles everything. When run on the unmodified original spreadsheets, it produces a spreadsheet equivalent to that seen in the instructions screenshots, the sole exception being that it also conditionally formats the Percent Change column, per the grading scale.

![2018 summary screenshot](Multiple_year_stock_data_2018.png "2018")
![2019 summary screenshot](Multiple_year_stock_data_2019.png "2019")
![2020 summary screenshot](Multiple_year_stock_data_2020.png "2020")

## Notes

The script does not store any information in variables if it is not required for the output. The grading scale says it should, though, so I can tweak that and resubmit if necessary.

Two scripts have been submitted. Both produce the same output, but the optimized script can perform the task on the entire Multiple_year_stock_data workbook in ~53 seconds, compared to the unoptimized script's ~217 seconds.

Only one screenshot per year is provided, to show the results of the script. The number of expected screenshots was not given, but providing enough screenshots to show every bit of the results would 237 screenshots, which I figured was a bit over the top, especially since the script can be run to provide the results, directly, in much less time than it would take to compile all those screenshots.

### Sources

Any information on concepts not provided in class was gleaned from various online sources, including:

UBound - https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/ubound-function

ReDim - https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/redim-statement

Function - https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/function-statement

NumberFormat - https://learn.microsoft.com/en-us/office/vba/api/excel.range.numberformat

AutoFit - https://learn.microsoft.com/en-us/office/vba/api/excel.range.autofit

Conditional Formatting - https://www.automateexcel.com/vba/conditional-formatting/

Using the Trim function on a number that has been cast to a String to remove leading whitespace - https://techcommunity.microsoft.com/t5/excel/runtime-error-1004-method-range-of-object-global-failed/m-p/215590
