"""
Provides functionality for advanced simulation options

"""

from numpy.lib.histograms import histogram
from py_expression_eval import Parser
import re
import numpy as np
from pyxll import xl_app, plot

#from meta_variables
from meta_variables import distributions_dictionary, id_location, screen_freeze_disabled
from meta_variables import simulation_num, explainError, cell_data, histogram_bins
import matplotlib.pyplot as plt



def advanced_simulation_cell(control, current_sheet_name = None, cell_address=None, variable_dict={},
                                    simulation_num=simulation_num, first_call=False):
    """
    Given a complex formula this returns a simulation. Option to manually pass in the 
    cell address for recursive applications, and likewise for variable_dict

    (variable_dict needs to be defined for recursive applications so we don't simulate the
    same cells multiple times)
    """
    xl = xl_app()

    #for recursion, have option to instead manually pass in an address
    if cell_address == None:
        cell_address = xl.Selection.Address
    else:
        cell_address=cell_address

    #Again for recursion, we need to be able to keep track of the current sheet we are on,
    #which may not be the sheet the user has selected
    if current_sheet_name == None:
        current_sheet_name = xl.ActiveSheet.Name
    else:
        current_sheet_name = current_sheet_name


    if re.search("[:,]", cell_address):
        """
        to check if the user has selected multiple cells by mistake
        """
        explainError(control=control, error_id="MultCellSelEr")
    else:
        pass

    formula = xl.Worksheets(current_sheet_name).Range(cell_address).Formula

    base_items = set(re.findall(r"[a-zA-Z0-9]{0,}[!]{0,1}[A-Z]+[0-9]+",formula))

    for curr_cell_address in base_items:
        if "!" in curr_cell_address:
            """
            this deals with the case where the address is a reference to another sheet

            e.g. Sheet2!A3 references A3 in Sheet2
            """
            sheet_name = curr_cell_address.split("!")[0]
            curr_cell_address = curr_cell_address.split("!")[1]
        else:
            sheet_name = current_sheet_name
            curr_cell_address = curr_cell_address

        transformed_cell_address = re.sub(r"\$[A-Z]+\$[0-9]+", lambda x: x.group()[1] + x.group()[3], curr_cell_address)
        #the above turns $A$1 into A1, and $AZ$36 into AZ36 etc.


        #e.g. Sheet1!A1 is stored as Sheet1___A1
        #If we are already in Sheet1, then cell A1 is stored as Sheet1___A1 in the dictionary
        sheet_cell_address = sheet_name + "_"*3 + transformed_cell_address

        cell_information = cell_data(control=control, cell_location=curr_cell_address, sheet_name=sheet_name)
        if cell_information != None:
            """covers the case where the cell is a distribution"""
            params = cell_information["params"] + [simulation_num]
            distr_id = cell_information["distribution_id"]

            if sheet_cell_address not in variable_dict:
                variable_dict[sheet_cell_address] = np.array(distributions_dictionary[distr_id]["scipy_handle"].rvs(*params))
            else:
                #in recursive applications, we might already have simulated from the cell
                pass
            
            #the above just gets the sample from the right distribution as a numpy array
        else:
            """Covers the case where the cell is another formula, or a constant"""
            index = re.search(r'[0-9]', curr_cell_address).start() 
            excel_address = "$" + curr_cell_address[:index] + "$" + curr_cell_address[index:]
            #we need to convert the address of the form say, A1, into $A$1 for being read
            #by excel when passed into advanced_simulation_cell as a address in excel


            """
            There's now three options. 
            0. Both the cell and its distribution data is empty
            1. the cell contains a fixed number, not a distribution
            2. the cell contains another formula
            3. the cell has a mistake
            """
            cell_value = xl.Worksheets[sheet_name].Range(excel_address).Value

            #case 0
            if cell_value == None:
                variable_dict[sheet_cell_address] = np.zeros(simulation_num)
                explainError(control=control, error_id="FormulaError",
                        custom_text=f"Cell {curr_cell_address}, sheet {sheet_cell_address} has no entry, default value of 0 used")
            
            #case 1 
            elif type(cell_value) == float or type(cell_value) == int:
                variable_dict[sheet_cell_address] = np.full((1, simulation_num), cell_value)

            #case 2. This is incomplete -  a formula could be wrong
            else:
                
                variable_dict[sheet_cell_address] = advanced_simulation_cell(control=control,
                                                    cell_address=excel_address,
                                                    variable_dict=variable_dict,
                                                    current_sheet_name=sheet_name)



    parser = Parser()

    #this adjusts the formula so it can be read and matched with our dictionary
    
    def g(x, formula_sheet_name=current_sheet_name):
        """
        Note that current_sheet_name refers to the sheet_name the formula is actually on
        whereas sheet_name at this point in the process refers to the sheet the last
        item we iterated through was on.

        Hence we want to turn something like Sheet1!$A$1 to Sheet1___A1
        and we want to turn something like $A$1 into <current_sheet_name>___A1
        where <current_sheet_name> is the name of the sheet the formula is on    
        """
        x = x.group()
        if "!" in x:
            y= x.split("!")
            return y[0]+"_"*3+y[1]
        else:
            return formula_sheet_name + "_"*3 + x
            
    formula = re.sub(r"[a-zA-Z0-9]{0,}[!]{0,1}[A-Z]+[0-9]+",g,formula)

    result = parser.parse(formula[1:]).evaluate(variable_dict)

    if first_call:
        #NEED TO LOOK INTO THIS
        variable_dict.clear()
    else:
        pass

    return result


def advanced_simulation_cell_wrapper(control, histogram_bins=histogram_bins):
    hist_data = advanced_simulation_cell(control=control, variable_dict={}), first_call=True

    fig, ax = plt.subplots(1, 1)

    ax.hist(hist_data, bins=histogram_bins)
    ax.grid()

    return plot(fig)   

