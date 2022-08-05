Project Title: Amazon Inventory Checker & Adjuster

Project Description: 
Amazon Vendor Central is a program which facilitates the sale of tools between customers and vendors. Each week Amazon sends vendors an excel sheet with all of the items they are requesting to purchase, and the orders are then confirmed or cancelled by the vendor, then shipped to fulfillment centers accross North America.

This program sorts the excel sheet order data provided, checking which orders are over $380 in value, then providing the total inventory requested for each product from these orders; this is useful as it makes checking/confirming inventory easier. The amount $380 is the minimum value threshold set, as orders under this amount are not worth the cost to package and ship (for the company using this program).

The second part of this program allows you to input the stock available for each inventory item requested - note: if all items are in stock then this step is skipped. The program then removes out-of-stock units from orders in a way that allows the most amount of orders to be fulfilled (based on the $380 value minimum).

Example: 
Order A: $1000 Value - Requesting 10 units of X @ $100/ea
Order B: $400 Value - Requesting 4 units of X @ $100/ea
Requested Units of X: 14
Actual Stock of X: 13
Result: this program will remove 1 unit of X from Order A, allowing both orders to be confirmed ($1300 total revenue).
Incorrect Result: this program will remove 1 unit of X from Order B, causing Order B to be cancelled ($1000 total revenue). 


How to Install and Run the Project:
Download Python (3.10.0)

How to Use the Project: 

Tests: 


Credits: 
Michelle Flandin


9 x CM100
20 x CT6100




Video Table of Contents: 
Purpose of the program
Program demonstration
Technologies used
How to install and run the program
How to use the program
How it works
In depth review of code



This program can be used by Amazon Vendor Central vendors to quickly evaluate the total amount of inventory being requested by Amazon in their weekly order spreadsheet, and if items are out of stock, this program will calculate the most efficient way to distribute the inventory among the purchase orders so that the maximum amount of profit can be earned. 


WHAT IS AMAZON VENDOR CENTRAL
Amazon Vendor Central is a web interface used by manufacturers who sell their products in bulk to Amazon for distribution purposes. Each week Amazon sends vendors an excel sheet with all of the items they are requesting to purchase, and the orders are then confirmed or cancelled by the vendor. The confirmed order are then shipped to Amazon fulfillment centers across North America, and the products are used to fulfill individual customer orders.


PROGRAM DEMO
Each row on the sheet, corresponds to a line item on a Purchase Order.
For example: 
Here on row 2 we have PO 5ZWYOJTT requesting 20 units of the GTC063, and available for download separately, we have the PO 5ZWYOJTT which shows the full purchase order.

There is also more information provided like the ship-to location, expected date, and more, but the only column information that will be relevant to this program will be:
COLUMN A: PO number
COLUMN G: Model Number
COLUMN N: Quantity Requested
COLUMN O: Expected Quantity
COLUMN P: Unit Cost
COLUMN Q: Currency, because there is a second sheet provided for orders going to Amazon Vendor Central Canada, which occassionally places orders with this company as well.

Sometimes Amazon orders more than is able to be produced by the expected date, and so before confirming these orders, we must consult our inventory - but this can take some time. In this case, I would need to sort by model number, then tally up each quantity requested, then compare the numbers with the inventory sheet. 


Whereas, with this program, I can instantly see how much of each unit is being requested.

I can also see which orders are above $380. For this particular company, shipping an order below $380 is not worth the shipping costs, so any order below $380 should be automatically cancelled and the inventory requested from those orders is therefore irrelevant and not included in the Inventory Summary.


If all the items are in stock, then nothing more needs to be done, but sometimes there is not enough inventory to fulfill these orders.

If this is the case, then we can open up the new spreadsheet created by this program.


Here we automatically see the second tab, and can input the number of units in stock 

Say we have only 20/40 units of the CT6100 available, and 0 units of the GTC062 available, we iput this data, save the file, close it, then run the second program: Amazon inventory adjuster.


Once it has run, it outputs the new orders to confirm and the inventory required for these orders. 

We can also re-open the file to see more details, which will make it easier when it comes to actually confirming the orders. 


We can also see the Inventory to Confirm sheet is updated here, with the Number of Units in Stock and from POs over $380

You may also have noticed the PO Raw Data sheet, which is just the copied data from the US and CAD orders all in one file. This is how the data was manipulated in the program, first creating a dictionary of each model number and the corresponding quantity requested, then putting that data into the newly created Inventory to Confirm sheet. The POs to Confirm sheet is also created, just to display the data more clearly, but the most important information for the first step is outputted to the console for ease of use. When Column D is updated in the inventory to confirm sheet on step 2, the data is then received and used to recalculate which orders need to be fulfilled. First it checks if the number of units is 0 - this is easy because if the units are zero, we simply delete all the line itmes with that model number, then recalculate the Purchase orders over $380. 

However, if there are some units in stock, but not all, then we need to first check if the units can be removed from any purchase orders without putting the total order value under $380. If this is not possible, then we start removing units from the lowest cost value orders, so that the most amount of revenue is achieved.



IN DEPTH CODE REVIEW:
First we have the main fucntion, which calls the pos_to_confirm file.

This file starts by checking to see if the sheets exist, as usually there are 2 sheets provided - 1 for the CAD orders and 1 for the USD orders, but sometimes there are only orders from the US or just from Canada.


# GO OVER CODE COMMENTS AND SEE IF YOU CAN WING IT WITH A FEW TRIALS

# RECORD VOICE (draft - be picky, but not too picky as it might change)

# RECORD + CUT THE VIDEO




