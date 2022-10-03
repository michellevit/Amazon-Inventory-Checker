Project Title: Amazon Inventory Checker & Adjuster


----------
Table of Contents: 
1. Project Summary
2. Detailed Overview
3. Example
4. Technologies Used
5. How to Install and Run the Script
6. How to Use the Program
7. Optional Changes
8. Troubleshooting
9. Credits


----------
1. Project Summary: 

This script is intended to automate the Amazon Vendor Central order confirmation process - it creates a concise requested inventory summary, with which the user can input the actual available inventory and instantly generate the final report needed for upload. This automation saves 5-20 minutes weekly and limits the potential for user error by decreasing the amount of manual data entry.


----------
2. Detailed Overview: 

Amazon Vendor Central is a website that deals with vendors selling bulk items to Amazon. Each week Amazon sends vendors an excel sheet with all of the items they are requesting to purchase, and the orders are then confirmed or cancelled by the vendor, then shipped to fulfillment centers accross North America.

This script was made for a vendor who has 2 conditions for accepting orders: 
1) The order is over $380 (if not, it is not worth accepting due to the shipping costs)
2) There is enough inventory

There are 2 python scripts - the first script intakes the weekly excel sheet, checks which orders are over $380, then outputs the inventory for each of the remaining orders. It also creates a new excel sheet with all the purchase orders that fulfill the first condition, and a separate sheet with the requested inventory for each item.

The second part of the script is employed if there is not enough inventory to complete the orders (which have passed the first condition). In this case, the user can open the newly created spreadsheet and enter the actual quantity of inventory that is available for each item. Then the user can run the second script, and the script will calculate which orders to cancel and output the results. In some cases there will be orders that can only be partially fulfilled, and this will be reflected in the new spreadsheet's final result. 

This script calculates the orders to cancel with an algorithm that aims to retain the most amount of profit, by first removing cancelled inventory 1-by-1 from orders which will not be put under the $380 threshold (thereby cancelling the entire order), then it removes cancelled inventory from orders starting from the lowest value to the highest value.


----------
3. Example: 
Order A: $1000 Value - Requesting 10 units of X @ $100/ea
Order B: $400 Value - Requesting 4 units of X @ $100/ea
Requested Units of X: 14
Actual Stock of X: 13

Expected/correct result: this program will remove 1 unit of X from Order A, allowing both orders to be confirmed ($1300 total revenue).
Incorrect result: this program will remove 1 unit of X from Order B, causing Order B to be cancelled ($1000 total revenue). 


----------
4. Technologies Used:
-Python (3.10.0)


----------
5. How to Install and Run the Script:
-Download Python (3.10.0)
-Double-click run_amazon_script_1.bat to run the first part of the script
-Double-click run_amazon_script_2.bat to run the second part of the script (must be done after running the first script)


----------
6. How to Use the Program:
-If there are US orders: save the Amazon line item file from Amazon Vendor Central US to the same folder as the scripts and .bat files
-If there are Canadian orders: save the Amazon line item file from Amazon Vendor Central US to the same folder as the scripts and .bat files
-Double-click the 'run_amazon_script_1.bat' file
-Open the newly created file 'POs to Confirm [month] [day], [year].xlsx"
-Confirm the inventory in the "Inv to Confirm' tab
-If there is not enough inventory: 
--Edit Column E (i.e. "Number of Units: In Stock and from POs > $380") in the 'Inv to Confirm' tab (highlighted yellow)
--Close the file
--Double-click the 'run_amazon_script_2.bat' file
--Open the 'POs to Confirm' tab
--Use this information to confirm the orders/inventory on Amazon Vendor Central


----------
7. Optional Changes: 
-To change the minimum PO value threshold amount: open the file 'amazon_script_1.py' and edit the global variable 'min_po_value' (line 12)


----------
8. Troubleshooting: 

-the excel file created, after running the first .bat file, must be closed before running the second .bat file
-the second .bat file cannot be run before the first
-Ensure the Amazon-generated files have the correct name(s) (i.e. 'EditLineItems.xlsx' and 'EditLineItems (1).xlsx')
-Ensure the Amazon-generated files have been downloaded to the correct folder

----------
9. Credits: 
Michelle Flandin
