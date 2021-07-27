#variables we will need
from meta_variables import distributions_dictionary, id_location, screen_freeze_disabled
from meta_variables import simulation_num, histogram_bins, error_messages_dictionary
from meta_variables import popupWindow_wrapper, DEBUG

from pyxll import xl_app, xl_func

def switchDebugMode(control, DEBUG=DEBUG, id_location=id_location):
    xl = xl_app()

    DEBUG["val"] = not DEBUG["val"]
    
    for sheet in xl.Sheets:
        #if a distribution sheet, we hide it if the distribution is set to false
        if xl.Worksheets(sheet.Name).Range(id_location).Value == "DISTRIBUTION SHEET":
            xl.Worksheets(sheet.Name).Visible = DEBUG["val"]
        else:
            pass

    return
