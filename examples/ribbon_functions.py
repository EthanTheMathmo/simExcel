"""
Defines the functions used in custom.xml, for the 'Add 1' button and the 'plot distribution' button 
"""
#general
from examples.meta_variables import cell_data
from pyxll import xl_app, xl_arg

#for distribution plotting
from examples.tkinter_frames import DistributionData
from pyxll import plot
import matplotlib.pyplot as plt
import numpy as np
import scipy.stats

#for custom user interface
from pyxll import xl_menu, create_ctp, CTPDockPositionFloating
import tkinter as tk
from tkinter_frames import DistributionData #I created tkinter_frames.py to hold tkinter code so this got less crowded

#for generating sheet names for the hidden sheets with distribution data
import random
import re

#variables we will need
from meta_variables import distributions_dictionary, id_location, screen_freeze_disabled
from meta_variables import simulation_num, histogram_bins, error_messages_dictionary
from meta_variables import popupWindow_wrapper

#Custom user interface
from pyxll import xl_menu, create_ctp, CTPDockPositionFloating
import tkinter as tk


"""
Initializing the distribution sheet, where we have a hidden sheet containing info for the cells
"""



def initSheetDistributionDict(control, id_location=id_location, screen_freeze_disabled=screen_freeze_disabled):
    """
    this initiates the hidden sheet with distribution info if it doesn't already exist.

    names are of the form 
        -SIM_ID
        -four digits
        -four lower case
        -four upper case
    e.g. %%!?1532afdkAFCT
    This format is so that 
    """
    xl = xl_app()

    upper_letters = [chr(n) for n in range(65,91)]
    lower_letters = [chr(n) for n in range(97,123)]
    digits = [str(i) for i in range(0,10)]

    name_for_sheet = "SIM_ID"+"".join(random.sample(digits,4))+"".join(random.sample(lower_letters,4)) +"".join(random.sample(upper_letters,4))

    if name_for_sheet not in [sheet.Name for sheet in xl.Sheets]:
        """
        this checks to make sure no names are duplicates
        """
        xl.ScreenUpdating = screen_freeze_disabled #if screen freezing is enabled
        #this freezes the screen so there is no flicker

        user_current_sheet_name = xl.ActiveSheet.Name #creating a new sheet will switch to the new sheet
        #but this sheet will be a hidden distribution sheet so we want to remember where we currently are
        xl.Sheets.Add()
        excel_current_sheet = xl.ActiveSheet
        excel_current_sheet.Name = name_for_sheet
        xl.Worksheets(user_current_sheet_name).Activate()

        cell_location = id_location
        xl.ActiveSheet.Unprotect() #if the cell we are writing to is protected then may need to unprotect it
        #this seems like a strictly worse solution than the hidden cell approach - should implement this later

        xl.ActiveSheet.Range(cell_location).Value = name_for_sheet 
        #NOTE later need to change cell_location to a different, hidden, cell as it will look ugly currently.
        xl.ActiveSheet.Range(cell_location).Locked = False #not sure if locking cell is sensible
        #xl.ActiveSheet.Protect() this would prevent other changes

        xl.ScreenUpdating = True
        return



    else:
        """
        If, by chance, the name is the same as another in the sheet, this runs the procedure again
        """
        initSheetDistributionDict(control=control)
        return

"""

This function is to extend the distribution functions so that it creates a 
distribution info sheet if one doesn't exist

"""

def distrInfoSheetInit(func, id_location=id_location):
    def wrapper(control, id_location=id_location):
        """
        control and id_location should be the names of two arguments of the input function
        """
        xl = xl_app()


        id_val = str(xl.ActiveSheet.Range(id_location).Value) #note if the cell is empty, this returns None, hence we need str for re.match to work
        if bool(re.match("SIM_ID[0-9]{4}[a-z]{4}[A-Z]{4}", id_val)):
            pass
        else:
            #initialize the new sheet if it doesn't
            initSheetDistributionDict(control=control)
            #note this function returns the ActiveSheet to the sheet the user was on when the function was initialized

        func(control=control, id_location=id_location)

    return wrapper
"""

INPUT DISTRIBUTIONS 

"""
#2 parameter distributions

@distrInfoSheetInit
def inputNormal(control, id_location = id_location):
    """
    This probably isn't the best way to do it, but when adding a button it seems to automatically pass in an argument
    so my work around is this wrapper function.
    """

    DistrInput(control=control, id_location=id_location, distr_id="N")

    return 


@distrInfoSheetInit
def inputExponential(control, id_location = id_location):

    DistrInput(control=control, id_location=id_location, distr_id="E")

    return 

@distrInfoSheetInit
def inputUniform(control, id_location=id_location):

    DistrInput(control=control, id_location=id_location, distr_id="U")

    return     

@distrInfoSheetInit
def inputSemicircular(control, id_location=id_location):

    DistrInput(control=control, id_location=id_location, distr_id="Semicircular")

    return  

@distrInfoSheetInit
def inputLaplace(control, id_location=id_location):

    DistrInput(control=control, id_location=id_location, distr_id="L")

    return  

@distrInfoSheetInit
def inputGumble_r(control, id_location=id_location):

    DistrInput(control=control, id_location=id_location, distr_id="Gumble_r")

    return  

@distrInfoSheetInit
def inputGumble_r(control, id_location=id_location):

    DistrInput(control=control, id_location=id_location, distr_id="Gumble_r")

    return  

#three parameter distributions

@distrInfoSheetInit
def inputTriangular(control, id_location = id_location):
    """
    This probably isn't the best way to do it, but when adding a button it seems to automatically pass in an argument
    so my work around is this wrapper function.
    """

    DistrInput(control=control, id_location=id_location, distr_id="T")

    return


@distrInfoSheetInit
def inputRice(control, id_location = id_location):
    """
    This probably isn't the best way to do it, but when adding a button it seems to automatically pass in an argument
    so my work around is this wrapper function.
    """

    DistrInput(control=control, id_location=id_location, distr_id="Rice")

    return

@distrInfoSheetInit
def inputPower_law(control, id_location = id_location):
    """
    This probably isn't the best way to do it, but when adding a button it seems to automatically pass in an argument
    so my work around is this wrapper function.
    """

    DistrInput(control=control, id_location=id_location, distr_id="Power_law")

    return

@distrInfoSheetInit
def inputPareto(control, id_location = id_location):
    """
    This probably isn't the best way to do it, but when adding a button it seems to automatically pass in an argument
    so my work around is this wrapper function.
    """

    DistrInput(control=control, id_location=id_location, distr_id="Pareto")

    return

@distrInfoSheetInit
def inputNakagami(control, id_location = id_location):
    """
    This probably isn't the best way to do it, but when adding a button it seems to automatically pass in an argument
    so my work around is this wrapper function.
    """

    DistrInput(control=control, id_location=id_location, distr_id="Nakagami")

    return


#four parameter distributions
@distrInfoSheetInit
def inputBetaPrime(control, id_location = id_location):
    """
    This probably isn't the best way to do it, but when adding a button it seems to automatically pass in an argument
    so my work around is this wrapper function.
    """

    DistrInput(control=control, id_location=id_location, distr_id="Bprime")

    return

@distrInfoSheetInit
def inputMielke(control, id_location = id_location):
    """
    This probably isn't the best way to do it, but when adding a button it seems to automatically pass in an argument
    so my work around is this wrapper function.
    """

    DistrInput(control=control, id_location=id_location, distr_id="Mielke")

    return


#the code for opening the distribution info window

def DistrInput(control, id_location, distr_id):
    """
    id_location is the cell containing the name of the sheet with distribution info
    """
    # Create the top level Tk window and give it a title
    window = tk.Toplevel()
    window.title("Input normal distribution data")

    # Create our example frame from the code above and add
    # it to the top level window.
    frame = DistributionData(master=window, control=control, screen_freeze_disabled=screen_freeze_disabled, id_location=id_location, distribution_id=distr_id)

    # Use PyXLL's 'create_ctp' function to create the custom task pane.
    # The width, height and position arguments are optional, but for this
    # example we'll create the CTP as a floating window rather than the
    # default of having it docked to the right.
    create_ctp(window,
               width=400,
               height=400,
               position=CTPDockPositionFloating
               )
    
    



"""

TEST BUTTON

"""

def on_text_button(control):
    xl = xl_app()
    xl.Selection.Value  += 1


"""

DISTRIBUTION INFOR

"""

def distribution_info(control, distributions_dictionary=distributions_dictionary):
    """
    not done yet. So that if someone clicks on a distribution key they get relevant info
    """
    xl = xl_app()
    cell_info = cell_data(control=control, cell_location=xl.Selection.Address)

    distr_id = cell_info["distribution_id"]

    params = [str(y) for y in cell_info["params"]]

    distr_name = distributions_dictionary[distr_id]["Name"]

    distr_params = distributions_dictionary[distr_id]["params"]

    title = "Distribution name: " + distr_name 

    text = "Distribution parameters: " + distr_params + "\n" + "Input parameters: " + ", ".join(params)

    popupWindow_wrapper(control=control, text=text, title=title)
    return




"""
PLOTTING DISTRIBUTIONS
"""



def display_distribution(control, screen_freeze_disabled=screen_freeze_disabled, 
                    id_location=id_location):
    """
    Input for this can be in two forms. Either
    (param1, param2,...) and (Dist_key) in a different cell

    Or, (param1, param2, ..., param_N, Dist_key)

    the distribution identifier in the cell should be in the distributions_dictionary defined at the top
    """
    xl = xl_app()

    user_sheet = xl.ActiveSheet.Name
    user_address = xl.Selection.Address

    if re.search("[,:]", user_address):
        """
        checks for if the user's selection is a single cell. If not then returns an error
        message
        """
        xl.Selection.Value = "MultCellSelEr"
        return
    else:
        pass

    distribution_sheet = xl.Range(id_location).Value
    #TO-DO # NOTE that we will need to add either an error message here if the distribtuion sheet doesn't exist,
    #or create it

    #retrieve the distribution data
    x = xl.Worksheets(distribution_sheet).Range(user_address).Value

    vals = x.split(",")
    params = [float(y) for y in vals[:-1]]
    dist_name = vals[-1]

    dist = distributions_dictionary[dist_name]["scipy_handle"](*params) #passes in our arguments as an array to the scipy function

    fig, ax = plt.subplots(1, 1)

    x = np.linspace(dist.ppf(0.01),
                    dist.ppf(0.99), 1000)
    ax.plot(x, dist.pdf(x),
        'r-', lw=2, alpha=0.6, label='norm pdf')
    ax.grid()

    return plot(fig)
    # try:
    #     num_params = distributions_dictionary[x[-1]]
    #     params = [float(y) for y in x[0].split(",")]
    #     if num_params != len(params):
    #         print("Incorrect number of parameters, or invalid values")
    #     else:

    # except KeyError:
    #     print("Distribution key not recognised") #need to learn how to raise an error in excel 



"""

SIMULATE the cells selected

"""
def parse_cell_sim(input_string, simulation_num, distributions_dictionary=distributions_dictionary):
    """
    Given a cell of the form (param1, param2, ..., distribution key),
    e.g. "0,1,N"
    and a number of simulations to perform, this returns 
    """
    y = input_string
    if input_string == None:
        return np.zeros(simulation_num)
    else:
        string_params = input_string.split(",")
        scipy_handle = distributions_dictionary[string_params[-1]]["scipy_handle"]
        numerical_params = [float(x) for x in string_params[:-1]] + [simulation_num]
        return np.array(scipy_handle.rvs(*numerical_params)) #for .rvs we pass in parameters, followed by size of sample

def parse_row_sim(input_row, simulation_num):
    """
    Given a row of data of the form
    (A,B,C,D...)
    where each element is a cell of the form passed into parse_cell_sim, this returns a numpy array of length simulation_num
    of the sums of elements in the row
    """
    np_array = np.array([parse_cell_sim(input_string=input_string, simulation_num=simulation_num) for input_string in input_row])
    return np_array.sum(axis=0)

def parse_block_sim(input_block, simulation_num):
    """
    If you've read the documentation for parse_cell_sim, and parse_row_sim, what this does should be 
    rather straightforward
    """
    np_array = np.array([parse_row_sim(input_row=input_row, simulation_num=simulation_num) for input_row in input_block])
    return np_array.sum(axis=0)

def hist_block_data(control, range_data, screen_freeze_disabled=screen_freeze_disabled,id_location=id_location, simulation_num=simulation_num, distributions_dictionary=distributions_dictionary):
    """
    turns a cell, row or block of data into the relevant simulation data (simple summing)
    """
    xl = xl_app()
    user_sheet = xl.ActiveSheet.Name

    #define range_data. For irregular data this is already defined, but for
    #if directed for regular_simulate the xl_app() isn't already defined so range_data is None
    if range_data == None:
        user_selection = xl.Selection.Address
    else:
        user_selection = range_data

    distribution_sheet = xl.ActiveSheet.Range(id_location).Value
    data = xl.Worksheets(distribution_sheet).Range(user_selection).Value

    if type(data) == str:
        hist_data = parse_cell_sim(data, simulation_num)
    else:
        if type(data[0]) == str:
            hist_data = parse_row_sim(data, simulation_num)
        else:
            hist_data = parse_block_sim(data, simulation_num)

    return hist_data
    
def regular_simulate(control, id_location=id_location, simulation_num=simulation_num, 
                distributions_dictionary=distributions_dictionary,
                histogram_bins = histogram_bins):
    """
    Specifically for regular shaped input (single cells, rows and blocks)
    """
    hist_data = hist_block_data(control=control, range_data = None, id_location=id_location, simulation_num=simulation_num, distributions_dictionary=distributions_dictionary)
    fig, ax = plt.subplots(1, 1)

    ax.hist(hist_data, bins=histogram_bins)
    ax.grid()

    return plot(fig)

def irregular_simulate(control, id_location=id_location, simulation_num=simulation_num, distributions_dictionary=distributions_dictionary):
    """
    Allows for irregular selection
    """
    xl = xl_app()
    user_selection = xl.Selection.Address
    user_sheet = xl.ActiveSheet.Name
    distribution_sheet = xl.ActiveSheet.Range(id_location).Value

    user_selection_array = user_selection.split(",")

    hist_data_notsummed = np.array([hist_block_data(control=control,
                            range_data=chunk,
                            id_location=id_location,
                            simulation_num=simulation_num) for chunk in user_selection_array])

    hist_data = hist_data_notsummed.sum(axis=0)
    fig, ax = plt.subplots(1, 1)


    ax.hist(hist_data)
    ax.grid()

    return plot(fig)


"""
Implement simulating a single simulation

"""


def cell_value_simulate(control, id_location=id_location):

    xl = xl_app()

    address = xl.Selection.Address

    single_simulation(control=control, address=address, id_location=id_location)

    return



def single_simulation(control, address, id_location=id_location):
    xl=xl_app()

    if "," in address:
        """
        this deals with the case where we have multiple blocks passed in by the user
        """
        chunks = address.split(",")
        for chunk in chunks:
            single_simulation(control, address=chunk, id_location=id_location)
    else:
        """
        this deals with the case where the user passes in a single block
        """
        xl.ActiveSheet.Range(address).Value = block_input_simulation(control=control, range_data=address,screen_freezing=screen_freeze_disabled, id_location=id_location)

    
from pyxll import xl_func

@xl_func
def one_cell_one_sim(input_string, distributions_dictionary=distributions_dictionary):
    """
    Given a cell of the form (param1, param2, ..., distribution key),
    e.g. "0,1,N"
    returns a single sample from the distribution
    """
    y = input_string
    if input_string == None:
        return "dist not specified"
    else:
        string_params = input_string.split(",")
        scipy_handle = distributions_dictionary[string_params[-1]]["scipy_handle"]
        numerical_params = [float(x) for x in string_params[:-1]] + [1]
        return scipy_handle.rvs(*numerical_params)[0] #for .rvs we pass in parameters, followed by size of sample

@xl_func
def one_row_one_sim(input_row):
    """
    Given a row of data of the form
    (A,B,C,D...)
    where each element is a cell of the form passed into parse_cell_sim, this returns a numpy array of length simulation_num
    of a single sample from each cell
    """
    x=[one_cell_one_sim(input_string=input_string) for input_string in input_row]
    return x

def one_block_one_sim(input_block):
    """
    If you've read the documentation for parse_cell_sim_1, and parse_row_sim, what this does should be 
    rather straightforward
    """
    return [one_row_one_sim(input_row=input_row) for input_row in input_block]

def block_input_simulation(control, range_data, screen_freezing=screen_freeze_disabled, id_location=id_location):
    """
    Given a cell block of inputs (where the selection object is of the form 
    $A$1
    or perhaps 
    $A$1:$H$6

    but NOT of the form $A$1:$H$6, $F$10, $W$1:$Z$10 we input a single simulation value for the cell

    This is used as a helper function for cell_value_simulate
    """
    xl = xl_app()

    user_sheet = xl.ActiveSheet.Name

    distribution_sheet = xl.ActiveSheet.Range(id_location).Value

    data = xl.Worksheets(distribution_sheet).Range(range_data).Value

    if type(data) == str:
        return_data = one_cell_one_sim(data)
    else:
        if type(data[0]) == str:
            return_data = one_row_one_sim(data)
        else:
            return_data = one_block_one_sim(data)

    return return_data





# def cell_value_simulate(control, id_location=id_location):
#     """
#     Given a selection of cells, runs a single simulation and inputs the value for each

#     This is the function used for the menu
#     """

#     xl = xl_app()
#     user_selection = xl.Selection.Address

    
#     user_selection_array = user_selection.split(",")

#     return_data = [block_input_simulation(control=control,
#                             range_data=chunk,
#                             id_location=id_location) for chunk in user_selection_array]
#     while True:
#         """
#         A somewhat hacky way of dealing with the nesting of array brackets which 
#         was occurring. So [[[[data]]]] -> [data], and if the data is a single entry, 
#         [[[[single_data_point]]]] -> single_data_point

#         Probably implement a tidier solution at some point
#         """
#         if len(return_data) == 1:
#             return_data = return_data[0]
#         else:
#             break

#     if "," in user_selection:
#         for index, chunk in enumerate(user_selection_array):
#             """
#             deals with the case where the user's input isn't a single cell or block
#             """
#             xl.ActiveSheet.Range(chunk).Value = return_data[index]
#     else:
#         xl.Selection.Range(user_selection).Value = return_data
#         xl.ActiveSheet.Range("$H$5").Value = user_selection

#     return