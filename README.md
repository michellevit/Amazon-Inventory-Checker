# Amazon Inventory Checker

### Date

April 2, 2022

### Assumptions Made

1. The input data should not be modified
2. There are less than 1000 rows of data

### Description

Each week, Amazon sends it's vendors an excel sheet which provides details on the orders it is requesting for that week.
The objective of this program is to find out the requested quantity per product, so that it is easy to consult the inventory lists and find out which product requests should be confirmed and which should be cancelled.
On the excel sheet, each row corresponds to a line item on a purchase order, providing information such as: the purchase order number, the product requested, the units requested, etc.
This program only uses the data from 2 columns: the product and it's requested quantity.

### Assumptions Made

1. The input data should not be modified

### Dependencies

1. Windows 10 (no other operating system tested)
2. Python library must be installed
3. openpyxl library must be installed

## Instructions / Executing the Program

1. Save the Amazon spreadsheet to the same folder/directory as this program
2. Open terminal
3. In the terminal, navigate to the directory that the program and Amazon excel sheet is saved in (using the cd command)
4. In the terminal, type: 'python amazon_order_tracker.py' to execute the script

### Tests/Debugging

There are 2 tests included in this program - both check that the column headers in the provided Amazon file match the column headers expected by the program.
If the tests fail, an 'AssertionError' will appear, and the program will terminate without providing results.

### Notes

The excel sheet included in this folder is a modified version of a file sent by Amazon - the modifications to the data are only meant to preserve the privacy of the vendor.

### Author

Michelle
