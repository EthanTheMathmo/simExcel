"""
This is for variables we use across several scripts

Also defines cell_data which retrieves the distribution info of a selected cell
"""
import scipy.stats
from pyxll import xl_app, xl_func
import re

distributions_dictionary = {"N":{"num_params":2, "scipy_handle":scipy.stats.norm, "params":"mean, variance","Name": "Normal Distribution"},
                            "C":{"num_params":2, "scipy_handle":scipy.stats.cauchy, "params": "mean, scaling","Name": "Cauchy"},
                            "T":{"num_params":3, "scipy_handle":scipy.stats.triang, "params": "c, loc, scale", "Name": "Triangular"},
                            "E":{"num_params":2, "scipy_handle":scipy.stats.expon, "params": "loc, scale", "Name":"Exponential"}}


id_location = "$A$1" #note the value of the id_location will
#at some point need to be changed to a hidden location
screen_freeze_disabled = True #for debugging, screen freezing often causes problems
#set to false to freeze screen while function operations are carried out

#number of simulations performed by default
simulation_num = 15000

#dictionary matching error codes to what the error is
PNumEr_str = "This error means you entered the wrong number of parameters for the distribution selected"
MultCellSelEr_str = "Multiple cells were selected and only one should have been"
error_messages_dictionary = {"PNumEr":PNumEr_str,
                            "MultCellSelEr":MultCellSelEr_str}


def cell_data(control, cell_location, id_location=id_location, 
            screen_freeze_disabled = screen_freeze_disabled):
    """
    Given a cell location, this returns the dictionary
    {"params"=[float, float, ....], "distribution_id": distribution_id}
    """
    
    xl = xl_app()

    distrInfoPageName = xl.ActiveSheet.Range(id_location).Value


    if re.search("[:,]", cell_location):
        """
        If the address for a block of cells has been passed in, this returns an error
        """
        xl.Selection.Value = "MultCellSelEr"
        return
    else:
        pass
    
    userCurrentPageName = xl.ActiveSheet.Name

    xl.ScreenUpdating = screen_freeze_disabled #this ensures no screen flickering from switching the active sheet

    xl.Worksheets(distrInfoPageName).Activate()

    #set the relevant values on the distrInfoSheet
    values = xl.ActiveSheet.Range(cell_location).Value.split(",")

    # "".join([form_result["Mean"],form_result["Standard deviation"], "N"])
    #return the active sheet to the user's original page
    xl.Worksheets(userCurrentPageName).Activate()

    xl.ScreenUpdating = True

    return_dict = {}
    return_dict["params"] = [float(val) for val in values[:-1]]
    return_dict["distribution_id"] = values[-1]

    return return_dict
    

