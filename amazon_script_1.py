import os, sys
from os import path
import openpyxl
from openpyxl import load_workbook
from openpyxl.styles.alignment import Alignment
from openpyxl.styles import PatternFill
from openpyxl.styles import Font, Alignment
from datetime import date
import operator

# GLOBAL VARIABLE 
min_po_value = 380

# sanity check
def pos_to_confirm():
    sheet_us = "Does not exist"
    sheet_ca = "Does not exist"
    if path.exists('EditLineItems.xlsx'):
        wb_unknown = load_workbook(filename = 'EditLineItems.xlsx')
        sheet_unknown = wb_unknown.active
        if sheet_unknown['Q2'].value == 'USD':
            wb_us = wb_unknown
            sheet_us = sheet_unknown
        elif sheet_unknown['Q2'].value == 'CAD':
            wb_ca = wb_unknown
            sheet_ca = sheet_unknown
        else:
            print("Note: filename 'EditLineItems.xlsx' does not exist.")
    if path.exists('EditLineItems (1).xlsx'):
        wb_unknown = load_workbook(filename = 'EditLineItems (1).xlsx')
        sheet_unknown = wb_unknown.active
        if sheet_unknown['Q2'].value == 'USD':
            wb_us = wb_unknown
            sheet_us = sheet_unknown
        elif sheet_unknown['Q2'].value == 'CAD':
            wb_ca = wb_unknown
            sheet_ca = sheet_unknown
        else: 
            print("Note: filename 'EditLineItems (1).xlsx' does not exist.")
    if (sheet_us == "Does not exist") and (sheet_ca == "Does not exist"): 
        print("Error: required files do not exist in directory (i.e. 'EditLineItems.xlsx and/or EditLineItems (1).xlsx)\nProgram cannot execute and will terminate.")
        exit()
    correct_wb(sheet_us)
    correct_wb(sheet_ca)
    populate_new_wb(sheet_us, sheet_ca, min_po_value)


# Test to make sure the file is compatible with this program, if it exists
def correct_wb(sheet):
    if (sheet == "Does not exist"):
        return
    else:
        assert sheet['A1'].value == 'PO',  'Cell A1 should be Model Number'
        assert sheet['G1'].value == 'Model Number',  'Cell G1 should be Model Number'
        assert sheet['N1'].value == 'Quantity Requested',  'Cell N1 should be Quantity Requested'
        assert sheet['O1'].value == 'Expected Quantity', 'Cell O1 should be Expected Quantity'
        assert sheet['P1'].value == 'Unit Cost', 'Cell P1 should be Unit Cost'


# Create a new excel excel workbook/file (at the end), containing the combined data from the US + CA orders (if they exist)
def populate_new_wb(sheet_us, sheet_ca, min_po_value):
    new_wb = openpyxl.Workbook()
    raw_data_sheet = new_wb.active
    raw_data_sheet.title = 'PO Raw Data'
    today = date.today()
    new_wb_filename = "POs to Confirm " + today.strftime("%B %d, %Y") + ".xlsx"
    # Copy raw data from sheet_us to new sheet
    if (sheet_us != "Does not exist"):
        for i in range(1, sheet_us.max_row + 1):
            for j in range(1, sheet_us.max_column + 1):
                raw_data_sheet.cell(row=i, column=j).value = sheet_us.cell(row=i, column=j).value
    # Append raw data from sheet_ca to new sheet
    if (sheet_ca != "Does not exist"):
        po_sheet_end = raw_data_sheet.max_row - 1
        for i in range(2, sheet_ca.max_row + 1):
            for j in range(1, sheet_ca.max_column + 1):
                raw_data_sheet.cell(row=i+po_sheet_end, column=j).value = sheet_ca.cell(row=i, column=j).value
    sorted_po_list = get_and_sort_po_values(raw_data_sheet, sheet_ca, min_po_value)
    create_inventory_tracker_dict(new_wb, raw_data_sheet, sorted_po_list, min_po_value)
    save_new_wb(new_wb, new_wb_filename)


# Create a dictionary with total value of each PO (key=PO, value=total PO value)
def get_and_sort_po_values(raw_data_sheet, sheet_ca, min_po_value):
    po_dict = {}
    # Calculate each POs total value, by iterating through values in column A and U (i.e. PO and Unit Cost), and save the data to po_dict
    for row in raw_data_sheet.iter_rows(min_row=2, max_row=raw_data_sheet.max_row):
        po = row[0].value
        cost = round(row[14].value * row[15].value) # cost times quantity
        # If po in dictionary, update quantity
        if po in po_dict.keys():
            po_dict[po] = round(po_dict[po] + cost)
        # If po not in dictionary, add it to dictionary
        else:
            po_dict[po] = cost
    # Sort POs by value, in descending order
    sorted_po_list = sorted(po_dict.items(), key=lambda x: x[1], reverse=True)
    print_pos_to_confirm(sorted_po_list, sheet_ca, min_po_value)
    return sorted_po_list


# Print the sorted 'POs to Confirm/Cancel' dictionary to Console
def print_pos_to_confirm(sorted_po_list, sheet_ca, min_po_value):
    cad_po_list = []
    if (sheet_ca != "Does not exist"):
        for row in sheet_ca.iter_rows(min_row=2, max_row=(sheet_ca.max_row)):
            po = row[0].value
            if po not in cad_po_list:
                cad_po_list.append(po)
    print('---------------\nPOs to Confirm/Cancel (stock not yet confirmed)\n---------------')
    longest_po_name = 0
    for item in sorted_po_list:
        po_number = item[0]
        if len(po_number) > longest_po_name:
            longest_po_name = len(po_number)
    for item in sorted_po_list:
        po_number = item[0]
        po_value = item[1]
        spaces_needed = longest_po_name - len(po_number)
        extra_space = ""
        for space in range(spaces_needed):
            extra_space = extra_space + ' '
        # print(po_number, extra_space, ' : ', po_value, sep='')
        ## If the order is from Canada, format it with '- CA' at the end, or else add blank spaces to the end
        if (po_number in cad_po_list):
            po_formatted = po_number + extra_space + ' (CAD) : '
        elif len(cad_po_list) != 0: 
            po_formatted = po_number + extra_space + '       : '
        else:
            po_formatted = po_number + extra_space + ' : '
        ## If the order is above 380, print it out with no additional note
        orders_to_cancel_list = create_list_pos_to_cancel(sorted_po_list, min_po_value)     
        if po_number in orders_to_cancel_list:
            cost_formatted = str(po_value)
            spaces_needed = 5 - len(cost_formatted)
            for space in range(spaces_needed):
                cost_formatted = cost_formatted + ' '
            print(po_formatted, cost_formatted, ' - CANCEL', sep='')
        ## If order value is below the min_po_value, format it with a note to cancel it
        else: 
            print(po_formatted, po_value, sep='')
            

# Create list of POs to cancel
def create_list_pos_to_cancel(sorted_po_list, min_po_value):
    orders_to_cancel_list = []
    for key, value in  sorted_po_list:
        if (value <= min_po_value) and (key not in orders_to_cancel_list):
            orders_to_cancel_list.append(key)
    return orders_to_cancel_list


# Keep track of the unit quantity ordered for each item (key=model number, value=units ordered)
def create_inventory_tracker_dict(new_wb, raw_data_sheet, sorted_po_list, min_po_value):
    ## Create empty lists/dicts
    inventory_list = []
    inventory_requested_dict = {}
    inventory_over_min_dict = {} 
    inventory_cancelled_dict = {}
    orders_to_cancel_list = create_list_pos_to_cancel(sorted_po_list, min_po_value)
    ## Iterate through values in the 'Raw Data Sheet'
    for row in raw_data_sheet.iter_rows(min_row=2, max_row=(raw_data_sheet.max_row)):
        product = row[6].value
        po_number = row[0].value
        quantity = round((row[14].value))
        if product not in inventory_list:
            inventory_list.append(product)
        if product in inventory_requested_dict:
            inventory_requested_dict[product] = inventory_requested_dict.get(product) + quantity
        else:
            inventory_requested_dict[product] = quantity  
        if po_number not in orders_to_cancel_list:
            if product in inventory_over_min_dict.keys():
                inventory_over_min_dict[product] = inventory_over_min_dict.get(product) + quantity
            else:
                inventory_over_min_dict[product] = quantity
            if product in inventory_cancelled_dict.keys():
                inventory_cancelled_dict[product] = inventory_cancelled_dict.get(product)
            else:
                inventory_cancelled_dict[product] = 0
        else: # if the product IS IN the orders_to cancel_list
            if product not in inventory_over_min_dict.keys():
                inventory_over_min_dict[product] = 0
            if product in inventory_cancelled_dict.keys():
                inventory_cancelled_dict[product] = inventory_cancelled_dict.get(product) + quantity
            else:
                inventory_cancelled_dict[product] = quantity
    # Sort the dict and list alphabetically
    sorted_inventory_list = sorted(inventory_list)
    sorted_inventory_over_min_dict = dict(sorted(inventory_over_min_dict.items(), key=operator.itemgetter(1),reverse=True))
    print_inventory_to_confirm(sorted_inventory_over_min_dict)
    create_inventory_sheet(new_wb, sorted_inventory_list, inventory_requested_dict, inventory_over_min_dict, inventory_cancelled_dict)
    create_pos_to_confirm_unadjusted_sheet(new_wb, raw_data_sheet, sorted_po_list, min_po_value)


# Print the sorted 'Inventory Requested (from non-cancelled POs)' dictionary
def print_inventory_to_confirm(sorted_inventory_over_min_dict):
    print('---------------\nInventory Requested (from non-cancelled POs)\n---------------')
    longest_product_name = 0
    for key in sorted_inventory_over_min_dict.keys():
        product = key
        if len(product) > longest_product_name:
            longest_product_name = len(product)
    for key, value in sorted_inventory_over_min_dict.items():
            if value != 0:
                product = key
                spaces_needed = longest_product_name - len(product)
                extra_space = ''
                for space in range(spaces_needed):
                    extra_space = extra_space + ' '
                print(product, extra_space, ' : ', value, sep='')


# Create a new sheet (in new_wb), with the POs and Requested Units
def create_inventory_sheet(new_wb, sorted_inventory_list, inventory_requested_dict, inventory_over_min_dict, inventory_cancelled_dict):
    inventory_sheet = new_wb.create_sheet('Inv to Confirm')
    new_wb.active = new_wb['Inv to Confirm']
    inventory_sheet.cell(row=1, column=1).value = "Model Number"
    inventory_sheet.cell(row=1, column=2).value = "Number of Units: \n Requested by Amazon"
    inventory_sheet.cell(row=1, column=3).value = "Number of Units: \n from POs > $380"
    inventory_sheet.cell(row=1, column=4).value = "Number of Units: \n In Stock"
    inventory_sheet.cell(row=1, column=5).value = "Number of Units: \n In Stock and from POs > $380"
    # Format the sheet nicely
    inventory_sheet.row_dimensions[1].height = 30
    for row in inventory_sheet.iter_cols(min_row=1, max_row=1, min_col=1, max_col=5):
        for cell in row:
            x = cell.coordinate
            inventory_sheet[x].font = Font(bold=True)
            inventory_sheet.column_dimensions[x[0]].width = 30
            inventory_sheet['A1'].fill = PatternFill("solid", start_color="CCE5FF")
            inventory_sheet['B1'].fill = PatternFill("solid", start_color="CCE5FF")
            inventory_sheet['C1'].fill = PatternFill("solid", start_color="CCE5FF")
            inventory_sheet['D1'].fill = PatternFill("solid", start_color="FFCCE5")
            inventory_sheet['E1'].fill = PatternFill("solid", start_color="E8E8E8")
            for y in range(inventory_sheet.max_row +1):
                for z in range(inventory_sheet.max_row +1):
                    for a in range(1, 6):
                        inventory_sheet.cell(row=z+1, column=a).alignment = Alignment(horizontal="center", wrap_text=True, vertical='center')
    populate_inventory_sheet(inventory_sheet, sorted_inventory_list, inventory_requested_dict, inventory_over_min_dict, inventory_cancelled_dict)


# Add data to 'Inv to Confirm' sheet
def populate_inventory_sheet(inventory_sheet, sorted_inventory_list, inventory_requested_dict, inventory_over_min_dict, inventory_cancelled_dict):     
    for i in range(len(sorted_inventory_list)):
        inventory_sheet.cell(row=i+2, column=1).value = sorted_inventory_list[i]
        inventory_sheet.cell(row=i+2, column=2).value = inventory_requested_dict.get(sorted_inventory_list[i])
        inventory_sheet.cell(row=i+2, column=3).value = inventory_over_min_dict.get(sorted_inventory_list[i])
        inventory_sheet.cell(row=i+2, column=4).value = inventory_over_min_dict.get(sorted_inventory_list[i])
    for row in inventory_sheet.iter_cols(min_row=2, max_row=inventory_sheet.max_row, min_col=4, max_col=4):
        for cell in row:
            if cell.value is None:
                break
            elif cell.value == 0:
                cell.font = Font(color="FFFFFF")
                cell.fill = PatternFill("solid", start_color="FFFDD0")
            else:
                cell.fill = PatternFill("solid", start_color="FFFDD0")
                cell.font = Font(color="A9A9A9")


# Create new sheet, containing the POs to Confirm
def create_pos_to_confirm_unadjusted_sheet(new_wb, raw_data_sheet, sorted_po_list, min_po_value):
    unadjusted_pos_sheet = new_wb.create_sheet('POs to Confirm')
    unadjusted_pos_sheet.cell(row=1, column=1).value = "PO Number"
    unadjusted_pos_sheet.cell(row=1, column=2).value = "Accepted/Cancelled"
    unadjusted_pos_sheet.cell(row=1, column=3).value = "All Items Accepted?"
    unadjusted_pos_sheet.cell(row=1, column=4).value = "Model Number"
    unadjusted_pos_sheet.cell(row=1, column=5).value = "Requested Quantity"
    unadjusted_pos_sheet.cell(row=1, column=6).value = "Accepted Quantity"
    unadjusted_pos_sheet.cell(row=1, column=7).value = "Currency"
    # Format the sheet nicely
    for row in unadjusted_pos_sheet.iter_cols(min_row=1, max_row=1, min_col=1, max_col=7):
        for cell in row:
            cell.alignment = Alignment(horizontal="center")
            x = cell.coordinate
            unadjusted_pos_sheet[x].font = Font(bold=True)
            unadjusted_pos_sheet.column_dimensions[x[0]].width = 20
            cell.fill = PatternFill("solid", start_color="ade6d4")
            for z in range(raw_data_sheet.max_row):
                for a in range(1, 8):
                    unadjusted_pos_sheet.cell(row=z+1, column=a).alignment = Alignment(horizontal="center")
    populate_pos_to_confirm_unadjusted_sheet(raw_data_sheet, unadjusted_pos_sheet, sorted_po_list, min_po_value)


# Populate sheet 'POs to Confirm' with data()
def populate_pos_to_confirm_unadjusted_sheet(raw_data_sheet, unadjusted_pos_sheet, sorted_po_list, min_po_value):
    # Iterate through values in column A (PO), G (Model Number), O (Expected Quantity), P (Unit Cost)
    for row in raw_data_sheet.iter_rows(min_row=2, max_row=(raw_data_sheet.max_row)):
        po_number = row[0].value
        model_number = row[6].value
        expected_quantity = row[14].value
        row_number = row[0].row
        currency = row[16].value
        unadjusted_pos_sheet.cell(row=row_number, column=1).value = po_number
        unadjusted_pos_sheet.cell(row=row_number, column=3).font = Font(color="008000")
        unadjusted_pos_sheet.cell(row=row_number, column=4).value = model_number
        unadjusted_pos_sheet.cell(row=row_number, column=5).value = expected_quantity
        unadjusted_pos_sheet.cell(row=row_number, column=6).value = expected_quantity
        orders_to_cancel_list = create_list_pos_to_cancel(sorted_po_list, min_po_value)
        if po_number in orders_to_cancel_list:
            unadjusted_pos_sheet.cell(row=row_number, column=2).value = 'CANCEL'
            unadjusted_pos_sheet.cell(row=row_number, column=2).font = Font(color="FF0000")
        else: 
            unadjusted_pos_sheet.cell(row=row_number, column=2).value = 'ACCEPT'
            unadjusted_pos_sheet.cell(row=row_number, column=2).font = Font(color="008000")
            unadjusted_pos_sheet.cell(row=row_number, column=3).value = 'YES'
        if currency == "USD":
            unadjusted_pos_sheet.cell(row=row_number, column=7).value = 'USD'
            unadjusted_pos_sheet.cell(row=row_number, column=7).font = Font(color="0000FF")
        else:
            unadjusted_pos_sheet.cell(row=row_number, column=7).value = 'CAD'
            unadjusted_pos_sheet.cell(row=row_number, column=7).font = Font(color="FF0000")
    # Format 'POs to Confirm Sheet' -> alternate background color of POs in 'POs to Confirm'
    prev = ""
    curr_color = "FFFFFF"
    alt_color = "E8E8E8"
    saved_color = "FFFFFF"
    for row in unadjusted_pos_sheet.iter_rows(min_row=2, max_row=(raw_data_sheet.max_row)):
        for cell in row:
            po_number = row[0].value
            if prev == "":
                cell.fill = PatternFill("solid", start_color=curr_color)
            elif po_number == prev:
                cell.fill = PatternFill("solid", start_color=curr_color)
            else: 
                cell.fill = PatternFill("solid", start_color=alt_color)
                saved_color = curr_color
                curr_color = alt_color
                alt_color = saved_color
            prev = po_number
    


# Save the workbook
def save_new_wb(new_wb, new_wb_filename):       
    new_wb.save(new_wb_filename)

