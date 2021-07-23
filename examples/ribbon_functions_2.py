"""
This is a second py file for the ribbons function because there was 
a certain degree of clutter


"""
#general
from pyxll import xl_app, xl_arg
import numpy as np








def parse_cell_sim_1(input_string):
    """
    Given a cell of the form (param1, param2, ..., distribution key),
    e.g. "0,1,N"
    returns a single sample, returning 0 if no entry currently exists
    """
    from ribbon_functions import distributions_dictionary
    y = input_string
    if input_string == None:
        return 0
    else:
        string_params = input_string.split(",")
        scipy_handle = distributions_dictionary[string_params[-1]]["scipy_handle"]
        numerical_params = [float(x) for x in string_params[:-1]] + [1]
        return scipy_handle.rvs(*numerical_params) #for .rvs we pass in parameters, followed by size of sample

def parse_row_sim_1(input_row):
    """
    Given a row of data of the form
    (A,B,C,D...)
    where each element is a cell of the form passed into parse_cell_sim, this returns a numpy array of length simulation_num
    of a single sample from each cell
    """
    return np.array([parse_cell_sim_1(input_string=input_string) for input_string in input_row])
    
def parse_block_sim_1(input_block):
    """
    If you've read the documentation for parse_cell_sim_1, and parse_row_sim, what this does should be 
    rather straightforward
    """
    return np.array([parse_row_sim_1(input_row=input_row) for input_row in input_block])

def block_input_simulation(control):
    """
    Given a cell block of inputs (where the selection object is of the form 
    $A$1
    or perhaps 
    $A$1:$H$6

    but NOT of the form $A$1:$H$6, $F$10, $W$1:$Z$10 we input a single simulation value for the cell

    This is used as a helper function for cell_value_simulate
    """
    xl = xl_app()

    from ribbon_functions import id_location

    user_sheet = xl.ActiveSheet.Name


    user_selection = xl.Selection.Address


    distribution_sheet = xl.ActiveSheet.Range(id_location).Value

    xl.ScreenUpdating = False
    xl.Worksheets(distribution_sheet).Activate()

    data = xl.ActiveSheet.Range(user_selection).Value

    if type(data) == str:
        return_data = parse_cell_sim_1(data)
    else:
        if type(data[0]) == str:
            return_data = parse_row_sim_1(data)
        else:
            return_data = parse_block_sim_1(data)



    xl.Worksheets(user_sheet).Activate()
    xl.ScreenUpdating = True

    return return_data





def cell_value_simulate(control):
    """
    Given a selection of cells, runs a single simulation and inputs the value for each

    This is the function used for the menu
    """
    from ribbon_functions import id_location

    xl = xl_app()
    user_selection = xl.Selection.Address

    user_selection_array = user_selection.split(",")

    return_data = np.array([block_input_simulation(control=control,
                            range_data=chunk,
                            id_location=id_location) for chunk in user_selection_array])

    xl.Selection.Value = return_data

    return