"""
Provides functionality for advanced simulation options

"""

from py_expression_eval import Parser
import re
import numpy as np
from pyxll import xl_app, plot
from meta_variables import distributions_dictionary, id_location, screen_freeze_disabled, simulation_num, explainError, cell_data
import matplotlib.pyplot as plt

def advanced_simulation_cell(control):
    """
    Given a complex formula this returns a simulation
    """
    xl = xl_app()

    cell_address = xl.Selection.Address

    if re.search("[:,]", cell_address):
        """
        to check if the user has selected multiple cells by mistake
        """
        explainError(control=control, error_id="MultCellSelEr")
    else:
        pass

    formula = xl.Selection.Formula

    base_items = set(re.findall('[A-Z]+[1-9]+',formula))

    variable_dict = {}

    for cell_address in base_items:
        transformed_cell_address = re.sub(r"\$[A-Z]+\$[1-9]+", lambda x: x.group()[1] + x.group()[3], cell_address)
        #the above turns $A$1 into A1, and $AZ$36 into AZ36 etc.
        cell_information = cell_data(control=control, cell_location=cell_address)
        params = cell_information["params"] + [simulation_num]
        distr_id = cell_information["distribution_id"]
        variable_dict[transformed_cell_address] = np.array(distributions_dictionary[distr_id]["scipy_handle"].rvs(*params))
        #the above just gets the sample from the right distribution as a numpy array

    parser = Parser()

    variable_dict["test"] = 333

    result = parser.parse(formula[1:]).evaluate(variable_dict)


    fig, ax = plt.subplots(1, 1)

    ax.hist(result)
    ax.grid()

    return plot(fig)


    

