"""
This is for ribbon functions used for understanding error messages,
for user feature requests
"""

#helpful variables
from meta_variables import distributions_dictionary, id_location, screen_freeze_disabled
from meta_variables import error_messages_dictionary
from meta_variables import cell_data

from pyxll import xl_app

def default_values(control, distribution_id, selection, params):
    """
    Sets the selected cells equal to their default values
    
    """
    xl = xl_app()

    if distribution_id == "N":
        xl.ActiveSheet.Range(selection).Value = float(params[0])
    elif distribution_id == "T":
        xl.ActiveSheet.Range(selection).Value = (float(params[2])-float(params[1]))*float(params[0]) + float(params[1])
    elif distribution_id == "E":
        xl.ActiveSheet.Range(selection).Value = float(params[0])+float(params[1]) 
    else:
        xl.ActiveSheet.Range(selection).Value = "<Need to add default value for this distribution. Search for this error in tkinter_frames.py>"

    return

def default_values_wrapper(control, id_location=id_location):
    """
    Defines relevant variables when called from the ribbon and executes default_values()
    for each cell in the selection

    Note: we do each cell individually because there is no guarantee the blocks are the 
    same distribution
    """
    xl = xl_app()

    for current_cell in xl.Selection:
        data = cell_data(control=control, cell_location=current_cell.Address)

        distribution_id = data["distribution_id"]
        params = data["params"]

        default_values(control, distribution_id=distribution_id, selection=current_cell.Address, params=params)


