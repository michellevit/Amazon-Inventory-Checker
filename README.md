# Amazon Inventory Checker

[NOTE: New version with Tkinter GUI (in progress)](https://github.com/michellevit/Amazon-Inventory-Checker-App)


![Python](https://img.shields.io/badge/Python-3.10-f7dd68.svg)
![openpyxl](https://img.shields.io/badge/OpenPyXL-3.0.9-206e47.svg)

A script which intakes Amazon's weekly order requests, and output the total inventory requested. If the total inventory requested is not available, the user can input their actual inventory and their minimum order value, and then the script will cancel units from orders selectively in order to maximize the total profit.


<a href="https://youtu.be/lD-wTry930w?si=mQKalJ-FNDY_Rot3" target="_blank"><img src="https://img.shields.io/badge/YouTube-Demo-red?style=for-the-badge&logo=youtube&color=FF0000"></a>


## Table of Contents
- [Technologies Used](#technologies-used)
- [Project Overview](#project-overview)
- [Example](#example)
- [How To Install](#how-to-install)
- [How To Use](#how-to-use)
- [Optional Changes](#optional-changes)
- [Troubleshooting](#troubleshooting)
- [Credits](#credits)


## Technologies Used<a name="technologies-used"></a>
- Python (3.10.0)
- OpenPyXL Library


## Project Overview<a name="project-overview"></a>
#### What is Amazon Vendor Central: 
Amazon Vendor Central is a website that deals with vendors selling bulk items to Amazon. Each week Amazon sends vendors an excel sheet with all of the items they are requesting to purchase, and the orders are then confirmed or cancelled by the vendor, then shipped to fulfillment centers accross North America.

#### Who Was This Script Made For?
This script was made for a vendor who has 2 conditions for accepting orders: 
1) The order is over $380 (if not, it is not worth accepting due to the shipping costs)
2) There is enough inventory

#### How Does the Script Work?
There are 2 python scripts - the first script intakes the weekly excel sheet, checks which orders are over $380, then outputs the total inventory for each model number (of the orders above $380). It also creates a new excel workbook with one sheet to clearly display all the inventory requested/cancelled, and another sheet to clearly display all of the order data.

The second part of the script is employed, if there is not enough inventory to complete all the orders (which have passed the first condition i.e. above $380). In this case, the user can open the newly created spreadsheet, open the inventory sheet, and enter the actual quantity of inventory that is available for each item. Then the user can run the second script, and the script will calculate which orders to remove the units from (prioritizing retaining maximum order profit), and then update the 'Vendor Download' workbooks (provided by Amazon), which can then be uploaded to be processed by Amazon Vendor Central. Once the Amazon Vendor Central website accepts the file, the confirmed orders will be available for download.

Regarding the algorithm which chooses which orders to be cancelled, if there is insufficient inventory, the algorithm first checks if removing 1 unit will put the order under the minimum threshold (i.e. $380), and if not, it will remove the unit from the order, and then proceed to either remove another unit from the same order (if more units must be cancelled, and the order has more units requested in it), or move onto the next order. If there are still units which need to be cancelled after trying to remove them from all the orders, the algorithm will then start removing units from the lowest-value order to the highest-value order. This algorithm could be improved, in the scenario that multiple products have low inventory, by taking into account which combination of units should be cancelled instead of removing product inventory one-at-a-time.


## Example<a name="example"></a>
Order A: $1000 Value - Requesting 10 units of X @ $100/ea
Order B: $400 Value - Requesting 4 units of X @ $100/ea
Requested Units of X: 14
Actual Stock of X: 13

Expected/correct result: this program will remove 1 unit of X from Order A, allowing both orders to be confirmed ($1300 total revenue).
Incorrect result: this program will remove 1 unit of X from Order B, causing Order B to be cancelled ($1000 total revenue). 


## How To Install<a name="how-to-install"></a>
- Clone the repository from GitHub
  - In the terminal, naviagte to the folder you want to download the project to
  - Run `git clone https://github.com/michellevit/Amazon-Inventory-Checker.git`
- Download Python (3.10.0) to the system
  - Make sure that the Python path is in the system's Environment Variables
  - To check and add Python to PATH:
  - Search for 'Environment Variables' in Windows Search and select "Edit the system environment variables".
  - In the System Properties window, click on "Environment Variables".
  - Under "System variables", scroll and find the 'Path' variable, then select "Edit".
  - If the path to Python is not listed, add it. This will typically be something like:
    - `C:\Users\Michelle\AppData\Local\Programs\Python\Python39\` 
    - *Note: adjust the path according to your Python version and installation location
  - You will likely need to restart your system after adding Python to the env variables
- Create a virtual environment
  - Navigate to C;\Users\*Username*\.virtualenvs
  - In the terminal, run `python -m venv Amazon-Checker-Virtual-Env`
- Activate the virtual environment
  - In the terminal, run `C:\Users\*Username*\.virtualenvs\Amazon-Checker-Virtual-Env\Scripts\Activate.ps1`
- Download openpyxl to the virtual environment: 
  - In the terminal (with the virtual env active), run `pip install openpyxl`
  - Check to see if it was installed, run `pip show openpyxl`
- Update the 'run_amazon_script_1' file:
  - Update the 'call' line in the 'run_amazon_script_1' file with folder location of the virtual environment
  - Update the 'python' line in the 'run_amazon_script_1' file with folder location of the project you just cloned from GitHub
- Update the 'run_amazon_script_2' file:
  - Update the 'call' line in the 'run_amazon_script_1' file with folder location of the virtual environment
  - Update the 'python' line in the 'run_amazon_script_1' file with folder location of the project you just cloned from GitHub


## How To Use<a name="how-to-use"></a>
- If there are US orders: save the Vendor Download file provided by Amazon Vendor Central into the folder with the .bat files
- If there are USD orders: save the Vendor Download file provided by Amazon Vendor Central, into the folder with the .bat files, and as a .xlsx file
- If there are CAD orders: save the Vendor Download file provided by Amazon Vendor Central, into the folder with the .bat files, and as a .xlsx file
- Double-click the 'run_amazon_script_1.bat' file
- Open the newly created file 'POs to Confirm [month] [day], [year].xlsx"
- Confirm the inventory in the "Inv to Confirm' tab
- If there is enough inventory:
  - Submit the file(s) to Amazon Vendor Central
- If there is not enough inventory: 
  - Edit Column E (i.e. "Number of Units: In Stock and from POs > $380") in the 'Inv to Confirm' tab (highlighted yellow)
  - Close the file
  - Double-click the 'run_amazon_script_2.bat' file
  - Submit the file(s) to Amazon Vendor Central


## Optional Changes<a name="optional-changes"></a>
- To change the minimum PO value threshold amount: 
  - open the file 'amazon_script_1.py'
  - edit the global variable 'min_po_value'


## Troubleshooting<a name="troubleshooting"></a>
- The excel file created, after running the first .bat file, must be closed before running the second .bat file
- The second .bat file cannot be run before the first
- The Amazon Vendor Download files are originally type 'Excel 97-2003 Workbook (*.xls), but they must be resave as type 'Excel Workbook (*.xlsx)' (and NOT as type 'Strict Open XML Spreadhseet (*.xlsx)')
- Ensure the Amazon Vendor Download files have been downloaded to the correct folder


## Credits<a name="credits"></a>
Michelle Flandin
