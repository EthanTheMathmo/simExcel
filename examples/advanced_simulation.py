"""
Provides functionality for advanced simulation options

"""

from py_expression_eval import Parser
import re
import numpy as np
from pyxll import xl_app, plot
from meta_variables import distributions_dictionary, id_location, screen_freeze_disabled, simulation_num, explainError, cell_data
import matplotlib.pyplot as plt

def advanced_simulation_cell(control, cell_address=None, variable_dict={}):
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

    if re.search("[:,]", cell_address):
        """
        to check if the user has selected multiple cells by mistake
        """
        explainError(control=control, error_id="MultCellSelEr")
    else:
        pass

    formula = xl.ActiveSheet.Range(cell_address).Formula

    base_items = set(re.findall('[A-Z]+[0-9]+',formula))

    for curr_cell_address in base_items:
        transformed_cell_address = re.sub(r"\$[A-Z]+\$[1-9]+", lambda x: x.group()[1] + x.group()[3], curr_cell_address)
        #the above turns $A$1 into A1, and $AZ$36 into AZ36 etc.
        cell_information = cell_data(control=control, cell_location=curr_cell_address)
        if cell_information != None:
            """covers the case where the cell is a distribution"""
            params = cell_information["params"] + [simulation_num]
            distr_id = cell_information["distribution_id"]
            if transformed_cell_address not in variable_dict:
                variable_dict[transformed_cell_address] = np.array(distributions_dictionary[distr_id]["scipy_handle"].rvs(*params))
            else:
                #in recursive applications, we might already have simulated from the cell
                pass
            
            #the above just gets the sample from the right distribution as a numpy array
        else:
            """Covers the case where the cell is another formula"""
            index = re.search(r'[1-9]', curr_cell_address).start()
            excel_address = "$" + curr_cell_address[:index] + "$" + curr_cell_address[index:]
            #we need to convert the address of the form say, A1, into $A$1 for being read
            #by excel when passed into advanced_simulation_cell as a address in excel
            variable_dict[transformed_cell_address] = advanced_simulation_cell(control=control,
                                                cell_address=excel_address,
                                                variable_dict=variable_dict)




    parser = Parser()

    variable_dict["test"] = 333

    result = parser.parse(formula[1:]).evaluate(variable_dict)

    return result


def advanced_simulation_cell_wrapper(control):
    hist_data = advanced_simulation_cell(control)

    fig, ax = plt.subplots(1, 1)

    ax.hist(hist_data)
    ax.grid()

    return plot(fig)   

