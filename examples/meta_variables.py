"""
This is for variables we use across several scripts

Also defines cell_data which retrieves the distribution info of a selected cell

This script should *not* import others in the file, as it is imported into nearly all of them
(circular dependency)
"""
from sre_constants import error
from tkinter.constants import BOTH
import scipy.stats
from pyxll import xl_app, xl_func, create_ctp, CTPDockPositionFloating
import re
import tkinter as tk

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

#histogram_bins
histogram_bins = 150

#dictionary matching error codes to what the error is
PNumEr_str = """Parameter Number Error.
This error means you entered the wrong number of parameters\n for the distribution selected"""
MultCellSelEr_str = """Multiple Cell Selection Error
Multiple cells were selected and only one should have been"""
ErrorButtonEr_str = "Oops - you selected multiple cells while using the error button"

GenericEr_str = "Input not recognised - please try selecting again"
#oops is reserved for a user mistake using the error message button
error_messages_dictionary = {"PNumEr":PNumEr_str,
                            "MultCellSelEr":MultCellSelEr_str,
                            "Oops!": ErrorButtonEr_str,
                            "Generic": GenericEr_str}

#debug. This currently does nothing, but the aim is that in the future it controls what sort of error
#messages might appear
DEBUG = True


"""
Returns cell distribution information

"""


def cell_data(control, cell_location, id_location=id_location, 
            screen_freeze_disabled = screen_freeze_disabled,
            literal=False,
            sheet_name = None):
    """
    Given a cell location, this returns the dictionary
    {"params"=[float, float, ....], "distribution_id": distribution_id}

    if Literal is set to True, this returns the actual string value

    In general sheet_name = None because we will want to be working in the sheet
    the user selected, but sheet_name gives is the option to override that
    """
    
    xl = xl_app()

    if re.search("[:,]", cell_location):
        """
        If the address for a block of cells has been passed in, this returns an error
        """
        xl.Selection.Value = "MultCellSelEr"
        return
    else:
        pass
    

    xl.ScreenUpdating = screen_freeze_disabled #this ensures no screen flickering from switching the active sheet

    userCurrentPageName = xl.ActiveSheet.Name

    #if a value for sheet_name has been passed in, we switch the activesheet to the one
    #specified
    if sheet_name == None:
        pass
    else:
        xl.Worksheets(sheet_name).Activate()

    distrInfoPageName = xl.ActiveSheet.Range(id_location).Value

    xl.Worksheets(distrInfoPageName).Activate()

    #get the relevant values on the distrInfoSheet
    values = xl.ActiveSheet.Range(cell_location).Value

    if literal == True:
        pass
    else:
        if values == None:
            """
            empty cell returns none. (i.e. cell with no distribution)
            """
            xl.Worksheets(userCurrentPageName).Activate() #return to user's page
            xl.ScreenUpdating = True

            return None
        else:
            values = values.split(",")

            return_dict = {}
            return_dict["params"] = [float(val) for val in values[:-1]]
            return_dict["distribution_id"] = values[-1]

    # "".join([form_result["Mean"],form_result["Standard deviation"], "N"])
    #return the active sheet to the user's original page
    xl.Worksheets(userCurrentPageName).Activate()

    xl.ScreenUpdating = True

    if literal:
        return values
    else:   
        return return_dict


"""
Implementing error tkinter window for use elsewhere



"""
class ErrorFrame(tk.Frame):

    def __init__(self, master, error_id, error_messages_dictionary=error_messages_dictionary):
        super().__init__(master)
        self.error_id = error_id
        self.error_messages_dictionary = error_messages_dictionary

        self.initUI()


    def initUI(self):
        # allow the widget to take the full space of the root window
        self.pack(fill=tk.BOTH, expand=True)

        # Create a tk.Label control and place it using the 'grid' method
        self.label_value = tk.StringVar()
        self.label = tk.Label(self, textvar=self.label_value)
        self.label.grid(column=0, row=1, sticky="w")
        self.label_value.set(self.error_messages_dictionary[self.error_id])


        # Allow the first column in the grid to stretch horizontally
        self.columnconfigure(0, weight=1)
 

def explainError(control, error_id, error_messages_dictionary=error_messages_dictionary):
    """
    Given an error id pop up an explanation of what it means
    """
        # Create the top level Tk window and give it a title
    window = tk.Toplevel()
    window.title("Error id: "+error_id)

    # Create our example frame from the code above and add
    # it to the top level window.
    frame = ErrorFrame(master=window, error_id=error_id)

    # Use PyXLL's 'create_ctp' function to create the custom task pane.
    # The width, height and position arguments are optional, but for this
    # example we'll create the CTP as a floating window rather than the
    # default of having it docked to the right.
    create_ctp(window,
               width=800,
               height=400,
               position=CTPDockPositionFloating)

