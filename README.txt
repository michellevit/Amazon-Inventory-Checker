Project Title: Amazon Inventory Checker & Adjuster


----------
Table of Contents: 
1. Project Description
2. Example
3. Technologies Used
4. How to Install and Run the Script
5. How to Use the Program
6. Optional Changes
7. Troubleshooting
8. Credits


----------
1. Project Description: 

SUMMARY: 
To help quickly calculate the total inventory being requested, from orders over $380. Also checks the most efficient order(s) to cancel inventory from, if not all units are in stock.

IN-DEPTH OVERVIEW:
Amazon Vendor Central is a website that deals with vendors selling bulk items to Amazon. Each week Amazon sends vendors an excel sheet with all of the items they are requesting to purchase, and the orders are then confirmed or cancelled by the vendor, then shipped to fulfillment centers accross North America.

This script was made for a vendor who has 2 conditions for acception orders: 
1) The order is over $380 (if not, it is not worth accepting due to the shipping costs)
2) There is enough inventory

There are 2 python scripts - the first script intakes the weekly excel sheet, checks which orders are over $380, then outputs the total inventory for each model number (of the orders above $380). It also creates a new excel workbook with one sheet to clearly display all the inventory requested/cancelled, and another sheet to clearly display all of the order data.

The second part of the script is employed, if there is not enough inventory to complete all the orders (which have passed the first condition i.e. above $380). In this case, the user can open the newly created spreadsheet, open the inventory sheet, and enter the actual quantity of inventory that is available for each item. Then the user can run the second script, and the script will calculate which orders to remove the units from (prioritizing retaining maximum order profit), and then update the 'Vendor Download' workbooks (provided by Amazon), which can then be uploaded to be processed by Amazon Vendor Central. Once the Amazon Vendor Central website accepts the file, the confirmed orders will be available for download.

Regarding the algorithm which chooses which orders to be cancelled, if there is insufficient inventory, the algorithm first checks if removing 1 unit will put the order under the minimum threshold (i.e. $380), and if not, it will remove the unit from the order, and then proceed to either remove another unit from the same order (if more units must be cancelled, and the order has more units requested in it), or move onto the next order. If there are still units which need to be cancelled after trying to remove them from all the orders, the algorithm will then start removing units from the lowest-value order to the highest-value order. This algorithm could be improved, in the scenario that multiple products have low inventory, by taking into account which combination of units should be cancelled instead of removing product inventory one-at-a-time.


----------
2. Example: 
Order A: $1000 Value - Requesting 10 units of X @ $100/ea
Order B: $400 Value - Requesting 4 units of X @ $100/ea
Requested Units of X: 14
Actual Stock of X: 13

Expected/correct result: this program will remove 1 unit of X from Order A, allowing both orders to be confirmed ($1300 total revenue).
Incorrect result: this program will remove 1 unit of X from Order B, causing Order B to be cancelled ($1000 total revenue). 


----------
3. Technologies Used:
-Python (3.10.0)


----------
4. How to Install and Run the Script:
-Download Python (3.10.0)
-Double-click run_amazon_script_1.bat to run the first part of the script
-Double-click run_amazon_script_2.bat to run the second part of the script (must be done after running the first script)


----------
5. How to Use the Program:
-If there are US orders: save the Vendor Download file provided by Amazon Vendor Central into the folder with the .bat files
-If there are USD orders: save the Vendor Download file provided by Amazon Vendor Central, into the folder with the .bat files, and as a .xlsx file
-If there are CAD orders: save the Vendor Download file provided by Amazon Vendor Central, into the folder with the .bat files, and as a .xlsx file
-Double-click the 'run_amazon_script_1.bat' file
-Open the newly created file 'POs to Confirm [month] [day], [year].xlsx"
-Confirm the inventory in the "Inv to Confirm' tab
-If there is not enough inventory: 
--Edit Column E (i.e. "Number of Units: In Stock and from POs > $380") in the 'Inv to Confirm' tab (highlighted yellow)
--Close the file
--Double-click the 'run_amazon_script_2.bat' file
--Submit the files to Amazon Vendor Central


----------
6. Optional Changes: 
-To change the minimum PO value threshold amount: open the file 'amazon_script_1.py' and edit the global variable 'min_po_value' (line 12)


----------
7. Troubleshooting: 

-the excel file created, after running the first .bat file, must be closed before running the second .bat file
-the second .bat file cannot be run before the first
-the Amazon Vendor Download files must be saved as .xlsx (they come as .xls, which openpyxl does not accept)
-Ensure the Amazon Vendor Download files have been downloaded to the correct folder

----------
8. Credits: 
Michelle Flandin