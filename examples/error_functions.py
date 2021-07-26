"""
This is for functions to do with error handling

This will be imported by other modules, and ideally should remain only having 
a dependency on meta_variables.py
"""

#helpful variables
from meta_variables import distributions_dictionary, id_location, screen_freeze_disabled
from meta_variables import error_messages_dictionary
from meta_variables import cell_data
from meta_variables import explainError
from meta_variables import DEBUG

from pyxll import xl_app, xl_menu

import re

def check_input(control, function_key):
    """
    Given a function and a selection, returns true or false depending on whether the selection is
    appropriate for that function, and also returns an error key

    NOT COMPLETE
    """
    xl = xl_app()

    user_selection = xl.ActiveSheet.Selection

    if function_key == "advanced_simulation_cell":
        #first we check if
        if re.find("[,;]", user_selection):
            return False
        else:
            pass
    else:
        2


def default_values(control, distribution_id, selection, params):
    """
    Sets the selected cells equal to their default values. In general this is the mean,
    although it is currently the median for triangular distribution.

    Check the scipy documentation for each distribution to see which parameter refers
    to what thing on the distribution
    
    """
    xl = xl_app()

    params = [float(x) for x in params]
    xl.ActiveSheet.Range(selection).Value = distributions_dictionary[distribution_id]["scipy_handle"].moment(*([1]+params))

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

def explainErrorWrapper(control):
    """
    Given an error code, explains it
    """
    xl = xl_app()

    if re.search("[:,]", xl.Selection.Address):
        """the button should only be used on a single entry"""
        explainError(control=control, error_id="Oops!", error_messages_dictionary=error_messages_dictionary)
        return
    else:
        error_id = xl.Selection.Value
        explainError(control=control, error_id=error_id, error_messages_dictionary=error_messages_dictionary)
        return

