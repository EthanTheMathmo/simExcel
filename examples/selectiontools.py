#variables we will need
from meta_variables import distributions_dictionary, id_location, screen_freeze_disabled
from meta_variables import simulation_num, histogram_bins, error_messages_dictionary
from meta_variables import popupWindow_wrapper

from pyxll import xl_app, xl_func

def selectDistrCells(control, id_location=id_location, sheet_name=None):
    """
    This allows the user to highlight all cells with the same distribution defined
    """
    xl = xl_app()

    if sheet_name==None:
        sheet_name = xl.ActiveSheet.Name
    else:
        sheet_name = sheet_name

    my_range = xl.Selection.Address
    distr_sheet_name = xl.Worksheets(sheet_name).Range(id_location).Value
    return_range = []
    
    for cell in xl.Selection:
        curr_address = cell.Address
        if xl.Worksheets(distr_sheet_name).Range(curr_address).Value != None:
            return_range.append(curr_address)
        else:
            pass
    
    if len(return_range) == 0:
        #if no items in return range, there are no cells to select
        xl.Range("A1").Select()
        return
    else:
        pass

    non_empty_cell_addresses = xl.Range(",".join(return_range))

    non_empty_cell_addresses.Select()

    return

def deleteSelectCells(control, id_location=id_location, sheet_name=None):
    """
    This allows the user to highlight all cells with the same distribution defined
    """
    xl = xl_app()

    if sheet_name==None:
        sheet_name = xl.ActiveSheet.Name
    else:
        sheet_name = sheet_name

    my_range = xl.Selection.Address
    distr_sheet_name = xl.Worksheets(sheet_name).Range(id_location).Value
    
    for cell in xl.Selection:
        curr_address = cell.Address
        xl.Worksheets(distr_sheet_name).Range(curr_address).Value = None



    return

