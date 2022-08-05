from multiprocessing.sharedctypes import Value
from openpyxl import load_workbook
import os
from datetime import date
from amazon_script_1 import *
import operator


def inventory_adjuster():
    # Check if 'POs to Confirm [DATE].xlsx' file exists (needed to run this program)
    today = date.today()
    pos_to_confirm_filename = "POs to Confirm " + today.strftime("%B %d, %Y") + ".xlsx"
    # If true, open first sheet, and proceed to next function
    if os.path.exists(pos_to_confirm_filename):
        new_wb = load_workbook(filename = pos_to_confirm_filename)
        inventory_sheet, items_with_zero_stock_list, items_with_some_stock_dict = get_inventory_to_cancel(new_wb)
        raw_data_sheet = new_wb["PO Raw Data"]
        deleted_inventory_dict = {}
        temporary_sheet = create_temporary_sheet(new_wb, raw_data_sheet, items_with_zero_stock_list, items_with_some_stock_dict, min_po_value, inventory_sheet, deleted_inventory_dict)
        inventory_to_confirm_dict = create_updated_inventory_to_confirm_dict(temporary_sheet)
        update_inventory_to_confirm_sheet(inventory_sheet, inventory_to_confirm_dict)
        orders_to_cancel_list = print_pos_to_confirm_to_console(raw_data_sheet, temporary_sheet)
        print_inventory_to_confirm_to_console(inventory_to_confirm_dict)
        update_pos_to_confirm_sheet(new_wb, temporary_sheet, orders_to_cancel_list)
        # del new_wb['Temporary Sheet']
        # del new_wb['PO Raw Data']
        new_wb.active = new_wb['POs to Confirm']
    # If false, print error message
    else: 
        print("Error - Filename: ", pos_to_confirm_filename, "does not exist.")
    save_new_wb(new_wb, pos_to_confirm_filename)


def get_inventory_to_cancel(new_wb):
    # Get 'Inv to Confirm' Sheet (so that you can put data into 'units_requested' and 'units_in_stock'  dicts)
    inventory_sheet = new_wb["Inv to Confirm"]
    items_with_zero_stock_list = []
    items_with_some_stock_dict = {} #kind of confusing but this dict is of key:[model number], value[units to cancel]
    # Get data from 'Inv to Confirm' Sheet
    for row in inventory_sheet.iter_rows(min_row=2, max_row=inventory_sheet.max_row):
        if row[0].value == None:
            break
        else:
            model_number = row[0].value
            units_requested = row[1].value
            units_in_stock = row[3].value
            if (units_in_stock == 0):
                items_with_zero_stock_list.append(model_number)
            elif (units_in_stock < units_requested and units_in_stock != 0):
                units_to_cancel = units_requested - units_in_stock
                items_with_some_stock_dict[model_number] = units_to_cancel
    # Reformat sheet (to make more readable)
    inventory_sheet['C1'].fill = PatternFill("solid", start_color="CCE5FF")
    inventory_sheet['D1'].fill = PatternFill("solid", start_color="CCE5FF")
    inventory_sheet['E1'].fill = PatternFill("solid", start_color="C1E1C1")
    # Remove background color from cells in column D
    for row in inventory_sheet.iter_cols(min_row=2, max_row=inventory_sheet.max_row, min_col=4, max_col=4):
        for cell in row:
            cell.fill = PatternFill(fill_type=None)
            cell.font = Font(color="000000")
    # Add background color to cells in column E
    for row in inventory_sheet.iter_cols(min_row=2, max_row=inventory_sheet.max_row, min_col=5, max_col=5):
        for cell in row:
            if cell.value == None:
                cell.fill = PatternFill(fill_type=None)
            else:
                cell.fill = PatternFill("solid", start_color="FFFDD0")
    return(inventory_sheet, items_with_zero_stock_list, items_with_some_stock_dict)


# Create new sheet with line items from all the CAD+USD POs, but:
## Delete line items if stock is zero
## Delete line items if PO is under minimum threshold
def create_temporary_sheet(new_wb, raw_data_sheet, items_with_zero_stock_list, items_with_some_stock_dict, min_po_value, inventory_sheet, deleted_inventory_dict):
    new_wb.copy_worksheet(raw_data_sheet)
    temporary_sheet = new_wb['PO Raw Data Copy']
    temporary_sheet.title = "Temporary Sheet"  
    # Delete rows with model numbers which have zero inventory, then delete items from inventory_to_cancel_dict
    if len(items_with_zero_stock_list) != 0:
        deleted_inventory_dict = delete_rows_with_zero_stock(items_with_zero_stock_list, temporary_sheet, deleted_inventory_dict)
    # Create dictionary with PO values from remaining non-deleted items
    po_value_dict = create_po_value_dict(temporary_sheet)
    # Create POs to cancel list, based on minimum threshold amount
    pos_under_min_threshold_list = create_pos_to_cancel_list(min_po_value, po_value_dict)
    # Delete rows from POs which are under min threshold
    if len(pos_under_min_threshold_list) != 0:
        deleted_inventory_dict = delete_pos_under_min_threshold(pos_under_min_threshold_list, temporary_sheet, deleted_inventory_dict)
    # Remove already deleted inventory from items_with_some_stock_dict
    if deleted_inventory_dict != 0 and items_with_some_stock_dict != 0:
        items_with_some_stock_dict = update_items_with_some_stock_dict(deleted_inventory_dict, items_with_some_stock_dict)
    # If there is still inventory that needs to be cancelled, go through POs and delete the units one-by-one
    if len(items_with_some_stock_dict) != 0:
        delete_remaining_out_of_stock_units(items_with_some_stock_dict, po_value_dict, min_po_value, temporary_sheet, inventory_sheet, deleted_inventory_dict)
    # Find PO values of remaining line items on temporary_sheet
    po_value_dict = create_po_value_dict(temporary_sheet)
    # Create a new POs to cancel list, based on minimum threshold amount and all the new changes
    pos_to_cancel_list = create_pos_to_cancel_list(min_po_value, po_value_dict)
    # Delete rows from POs which are under min threshold
    if len(pos_to_cancel_list) != 0:
        deleted_inventory_dict = delete_pos_under_min_threshold(pos_under_min_threshold_list, temporary_sheet, deleted_inventory_dict)
    return(temporary_sheet)


def delete_rows_with_zero_stock(items_with_zero_stock_list, temporary_sheet, deleted_inventory_dict):
    rows = list(temporary_sheet.iter_rows(min_row=2, max_row=temporary_sheet.max_row))
    rows = reversed(rows)
    for row in rows:
        model_number = row[6].value
        requested_quantity = row[13].value
        if model_number in items_with_zero_stock_list:
            temporary_sheet.delete_rows(row[0].row, 1)
            if model_number in deleted_inventory_dict.keys():
                deleted_inventory_dict[model_number] = deleted_inventory_dict[model_number] + requested_quantity
            else:
                deleted_inventory_dict[model_number] = requested_quantity
    return(deleted_inventory_dict)


def create_po_value_dict(temporary_sheet):
    po_value_dict = {}
    for row in temporary_sheet.iter_rows(min_row=2, max_row=(temporary_sheet).max_row):
        po_number = row[0].value
        quantity = row[14].value
        unit_cost = round(row[15].value, 2)
        if po_number not in po_value_dict:
            po_value_dict[po_number] = round(quantity * unit_cost, 2)
        else:
            po_value_dict[po_number] = round(po_value_dict.get(po_number) + round((quantity * unit_cost), 2), 2)
    return(po_value_dict)


def create_pos_to_cancel_list(min_po_value, po_value_dict):
    pos_under_min_threshold_list = []
    for key, value in po_value_dict.items():
        if value < min_po_value:
            pos_under_min_threshold_list.append(key)
    return(pos_under_min_threshold_list)


def delete_pos_under_min_threshold(pos_under_min_threshold_list, temporary_sheet, deleted_inventory_dict):
    rows = list(temporary_sheet.iter_rows(min_row=2, max_row=temporary_sheet.max_row))
    rows = reversed(rows)
    for row in rows:
        po_number = row[0].value
        model_number = row[6].value
        requested_quantity = row[13].value
        if po_number in pos_under_min_threshold_list:
            temporary_sheet.delete_rows(row[0].row, 1)
            if model_number in deleted_inventory_dict.keys():
                deleted_inventory_dict[model_number] = deleted_inventory_dict[model_number] + requested_quantity
            else:
                deleted_inventory_dict[model_number] = requested_quantity
    return(deleted_inventory_dict)


def update_items_with_some_stock_dict(deleted_inventory_dict, items_with_some_stock_dict):
    for key, value in items_with_some_stock_dict.items():
        if key in deleted_inventory_dict.keys():
            items_with_some_stock_dict[key] = items_with_some_stock_dict[key] - deleted_inventory_dict[key]
    return(items_with_some_stock_dict)

    
def delete_remaining_out_of_stock_units(items_with_some_stock_dict, po_value_dict, min_po_value, temporary_sheet, inventory_sheet, deleted_inventory_dict):
    rows = list(temporary_sheet.iter_rows(min_row=2, max_row=temporary_sheet.max_row))
    rows = reversed(rows)
    for row in rows:
        model_number = row[6].value
        if model_number in items_with_some_stock_dict.keys():
            units_to_cancel = items_with_some_stock_dict[model_number]
            po_number = row[0].value
            unit_cost = row[15].value
            units_cancelled = 0    
            for i in range(units_to_cancel):
                if ((po_value_dict[po_number] - unit_cost) < min_po_value):
                    break
                else:   
                    if row[14].value == 0:
                        break
                    else: 
                        temporary_sheet.cell(row=row[0].row, column=15).value = row[14].value - 1
                        temporary_sheet.cell(row=row[0].row, column=15).fill = PatternFill("solid", start_color="FFFFE0")
                        units_cancelled = units_cancelled + 1
                        po_value_dict[po_number] = round(po_value_dict[po_number] - unit_cost, 2)
            items_with_some_stock_dict[model_number] = items_with_some_stock_dict[model_number] - units_cancelled
            if items_with_some_stock_dict[model_number] == 0:
                items_with_some_stock_dict.pop(model_number)
    # If there are still units to be cancelled, then remove them starting from the lowest value pos
    # Remove POs from dict, which are under the min value
    for key, value in list(po_value_dict.items()):
        if value < min_po_value:
            del po_value_dict[key]
    low_to_high_po_value_list = sorted(po_value_dict.items(), key=lambda x: x[1])
    ## Check if there is still stock to cancel
    if len(items_with_some_stock_dict) != 0:
        ## Iterate through the rows of the filtered line items, starting with the lowest value PO
        for index, tuple in enumerate(low_to_high_po_value_list):
            list_po_number = tuple[0]
            list_po_value = tuple[1]
            rows = list(temporary_sheet.iter_rows(min_row=2, max_row=temporary_sheet.max_row))
            rows = reversed(rows)
            for row in rows:
                if len(items_with_some_stock_dict) == 0:
                    break
                row_po_number = row[0].value
                if list_po_number == row_po_number:
                    row_model_number = row[6].value
                    units_cancelled = 0
                    if row_model_number in items_with_some_stock_dict.keys():
                        units_to_cancel = items_with_some_stock_dict[row_model_number]
                        for i in range(units_to_cancel):
                            if row[14].value == 0:
                                temporary_sheet.delete_rows(row[14].row, 1)
                                break
                            else: 
                                row[14].value = row[14].value - 1
                                units_cancelled = units_cancelled + 1
                        items_with_some_stock_dict[row_model_number] = items_with_some_stock_dict[row_model_number] - units_cancelled
                        for key, value in list(items_with_some_stock_dict.items()):
                            if value == 0:
                                del items_with_some_stock_dict[key]
    po_value_dict = create_po_value_dict(temporary_sheet)
    pos_to_cancel_list = create_pos_to_cancel_list(min_po_value, po_value_dict) 
    deleted_inventory_dict = delete_pos_under_min_threshold(pos_to_cancel_list, temporary_sheet, deleted_inventory_dict)            
    correct_if_too_much_cancelled(temporary_sheet, inventory_sheet)


def correct_if_too_much_cancelled(temporary_sheet, inventory_sheet): 
    tentatively_confirmed_units_dict = {}
    rows = list(temporary_sheet.iter_rows(min_row=2, max_row=temporary_sheet.max_row))
    for row in rows: 
        model_number = row[6].value
        expected_quantity = row[14].value
        if model_number in tentatively_confirmed_units_dict.keys():
            tentatively_confirmed_units_dict[model_number] = tentatively_confirmed_units_dict[model_number] + expected_quantity
        else:
            tentatively_confirmed_units_dict[model_number] = expected_quantity
    units_in_stock_dict = get_units_in_stock(inventory_sheet)
    for key in list(units_in_stock_dict):
        if key in tentatively_confirmed_units_dict.keys():
            units_in_stock_dict[key] = units_in_stock_dict[key] - tentatively_confirmed_units_dict[key]
    for key in list(units_in_stock_dict):
        if units_in_stock_dict[key] == 0:
            del units_in_stock_dict[key]
    units_to_salvage_dict = units_in_stock_dict
    rows = list(temporary_sheet.iter_rows(min_row=2, max_row=temporary_sheet.max_row))
    for row in rows:
        row_model_number = row[6].value
        if row_model_number in units_to_salvage_dict:
            quantity_requested = row[13].value
            expected_quantity = row[14].value
            added_amount = 0
            if expected_quantity < quantity_requested:
                for i in range(units_to_salvage_dict[row_model_number]):
                    row[14].value = row[14].value + 1
                    added_amount = added_amount + 1
                    if row[14].value == quantity_requested:
                        break                      
            units_to_salvage_dict[row_model_number] = units_to_salvage_dict[row_model_number] - added_amount
            if units_to_salvage_dict[row_model_number] == 0:
                del units_to_salvage_dict[key]

def get_units_in_stock(inventory_sheet):
    units_in_stock_dict = {}
    rows = list(inventory_sheet.iter_rows(min_row=2, max_row=inventory_sheet.max_row))
    for row in rows:
        if row[0].value == None:
            break
        else: 
            model_number = row[0].value
            units_in_stock = row[3].value
            if model_number in units_in_stock_dict.keys():
                units_in_stock_dict[model_number] = units_in_stock_dict[model_number] + units_in_stock
            else:
                units_in_stock_dict[model_number] = units_in_stock
    return(units_in_stock_dict)


def create_updated_inventory_to_confirm_dict(temporary_sheet):
    inventory_to_confirm_dict = {}
    rows = list(temporary_sheet.iter_rows(min_row=2, max_row=temporary_sheet.max_row))
    for row in rows:
        if row[0].value == None:
            break
        model_number = row[6].value
        quantity = (row[14].value)
        if model_number not in inventory_to_confirm_dict.keys():
            inventory_to_confirm_dict[model_number] = quantity
        else:
            inventory_to_confirm_dict[model_number] = inventory_to_confirm_dict[model_number] + quantity
    return(inventory_to_confirm_dict)    


def update_inventory_to_confirm_sheet(inventory_sheet, inventory_to_confirm_dict):
    for row in inventory_sheet.iter_rows(min_row=2, max_row=(inventory_sheet.max_row)):
        model_number = row[0].value
        if row[0].value == None:
            break
        elif row[0].value in inventory_to_confirm_dict:
            row[4].value = inventory_to_confirm_dict[model_number]
            row[4].fill = PatternFill("solid", start_color="FFFDD0")
        else:
            row[4].value = 0
            row[4].fill = PatternFill("solid", start_color="FFFDD0")


# Print the 'POs to Confirm/Cancel' to Console
def print_pos_to_confirm_to_console(raw_data_sheet, temporary_sheet):
    pos_to_confirm_list = []
    rows = list(temporary_sheet.iter_rows(min_row=2, max_row=temporary_sheet.max_row))
    for row in rows: 
        if row[0].value not in pos_to_confirm_list:
            pos_to_confirm_list.append(row[0].value)
    cad_po_list = []
    crows = list(raw_data_sheet.iter_rows(min_row=2, max_row=raw_data_sheet.max_row))
    for crow in crows:
        if (crow[16].value == "CAD") and (crow[0].value not in cad_po_list):
            cad_po_list.append(crow[0].value)
    po_value_dict = create_po_value_dict(temporary_sheet)
    high_to_low_po_value_list = sorted(po_value_dict.items(), key=lambda x: x[1], reverse=True)
    print('---------------\n*UPDATED* POs to Confirm/Cancel (From POs > 380 + In Stock)\n---------------')
    for index, tuple in enumerate(high_to_low_po_value_list):
        po_number = tuple[0]
        po_value = tuple[1]
        ## If the order is from Canada, format it with '- CA' at the end, or else add blank spaces to the end
        if po_number in cad_po_list:
            po_formatted = po_number + ' (CAD) : '
        elif len(cad_po_list) != 0: 
            po_formatted = po_number + '       : '
        else:
            po_formatted = po_number + ' : '
        print(po_formatted, po_value)
    orders_to_cancel_list = create_list_pos_to_cancel(raw_data_sheet, pos_to_confirm_list)
    for x in orders_to_cancel_list:
        print(x, "- CANCEL")
    return(orders_to_cancel_list)


def create_list_pos_to_cancel(raw_data_sheet, pos_to_confirm_list):
    orders_to_cancel_list = []
    rows = list(raw_data_sheet.iter_rows(min_row=2, max_row=raw_data_sheet.max_row))
    for row in rows: 
        if (row[0].value not in pos_to_confirm_list) and (row[0].value not in orders_to_cancel_list):
            orders_to_cancel_list.append(row[0].value)
    return(orders_to_cancel_list)


# Print the sorted 'POs to Confirm/Cancel' to Console
def print_inventory_to_confirm_to_console(inventory_to_confirm_dict):
    print('---------------\n*UPDATED* Inventory to Confirm (From POs > 380 + In Stock)\n---------------')
    longest_product_name = 0
    for key in inventory_to_confirm_dict.keys():
        product = key
        if len(product) > longest_product_name:
            longest_product_name = len(product)
    sorted_inventory_to_confirm_dict = dict(sorted(inventory_to_confirm_dict.items(), key=operator.itemgetter(1),reverse=True))
    for key, value in sorted_inventory_to_confirm_dict.items():
        product = key
        spaces_needed = longest_product_name - len(product)
        extra_space = ""
        for space in range(spaces_needed):
            extra_space = extra_space + ' '
        print(product, extra_space, ' : ', value, sep='')


def update_pos_to_confirm_sheet(new_wb, temporary_sheet, orders_to_cancel_list):
    pos_to_confirm_sheet = new_wb["POs to Confirm"]
    rows = list(pos_to_confirm_sheet.iter_rows(min_row=2, max_row=pos_to_confirm_sheet.max_row))
    for row in rows: 
        new_po_number = row[0].value
        new_model_number = row[3].value
        new_accepted_quantity = row[5].value
        checker = 0
        if new_po_number in orders_to_cancel_list:
            row[1].value = "CANCELLED"
            row[1].font = Font(color="FF0000")
            row[2].value = " "
            row[5].value = 0
        else: 
            crows = list(temporary_sheet.iter_rows(min_row=2, max_row=temporary_sheet.max_row))
            for crow in crows: 
                old_po_number = crow[0].value
                old_model_number = crow[6].value
                old_expected_quantity = crow[14].value
                if (new_po_number == old_po_number) and (new_model_number == old_model_number):
                    if new_accepted_quantity != old_expected_quantity:
                        row[5].value = old_expected_quantity
                        row[5].fill = PatternFill("solid", start_color="FFFDD0")
                        row[2].value = "NO"
                        row[2].font = Font(color="FF0000")
                        checker = 1
                    elif new_accepted_quantity == old_expected_quantity:
                        checker = 1
            if checker == 0:
                row[1].value = "CANCELLED"
                row[1].font = Font(color="FF0000")
                row[2].value = "-"
                row[5].value = 0
                row[5].font = Font(color="FF0000")
                row[5].value = 0


# Save the workbook
def save_new_wb(new_wb, pos_to_confirm_filename): 
    new_wb.save(pos_to_confirm_filename)



inventory_adjuster()
