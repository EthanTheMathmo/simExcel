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

#for encoding sheet name as integers
import sys

#for summary statistics
import scipy.stats


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

    set1 = set([x for x in re.findall(r"[a-zA-Z0-9]{0,}[!]{0,1}[A-Z]+[0-9]+",formula) if x[0]!="!"])

    set2 = set([x.replace("'", "") for x in re.findall(r"['].{1,}?['][!]{0,1}[A-Z]+[0-9]+",formula)])
    base_items = set1.union(set2)

    for curr_cell_address in base_items:
        if "!" in curr_cell_address:
            """
            this deals with the case where the address is a reference to another sheet

            e.g. Sheet2!A3 references A3 in Sheet2

            Note this may seem a little opague. This is because it must deal with the case
            where there are ! in the title. E.g. The sheet could be called
            Abb!!2321A, in which case to reference A5 we have

            Abb!!2321A!A5

            Hence we need to split of the bit after the *last* exclam, and the bit before
            """
            split_by_exclam = curr_cell_address.split("!")
            sheet_name = "!".join(split_by_exclam[:-1])
            curr_cell_address = split_by_exclam[-1]
        else:
            sheet_name = current_sheet_name
            curr_cell_address = curr_cell_address

        transformed_cell_address = re.sub(r"\$[A-Z]+\$[0-9]+", lambda x: x.group()[1] + x.group()[3], curr_cell_address)
        #the above turns $A$1 into A1, and $AZ$36 into AZ36 etc.


        #If we are already in Sheet1, cell A1 is stored as follows. Sheet1!A1 is stored as an integer encoding, and then we write the integer
        #as a string. The ! is needed to prevent two different cells having the same encoding. E.g. if we had a sheet named 
        #A and one named AA, then AA1 in A1 would be given the same encoding as A1 in AA.
        sheet_cell_encoding = "F" + str(int.from_bytes((sheet_name+"!"+transformed_cell_address).encode(encoding="utf-8"), byteorder = sys.byteorder))
        #sheet_cell_address = sheet_name + "_"*3 + transformed_cell_address THIS IS A CRAP SOLUTION

        cell_information = cell_data(control=control, cell_location=curr_cell_address, sheet_name=sheet_name)
        if cell_information != None:
            """covers the case where the cell is a distribution"""
            params = cell_information["params"] + [simulation_num]
            distr_id = cell_information["distribution_id"]

            if sheet_cell_encoding not in variable_dict:
                variable_dict[sheet_cell_encoding] = np.array(distributions_dictionary[distr_id]["scipy_handle"].rvs(*params))
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
            0. the cell contains another formula
            1. the cell contains a fixed number, not a distribution
            2. Both the cell and its distribution data is empty
            3. the cell has a mistake
            """
            cell_value = xl.Worksheets[sheet_name].Range(excel_address).Value

            cell_formula = xl.Worksheets(sheet_name).Range(excel_address).Formula

            #case 0
            if cell_formula != None:
                variable_dict[sheet_cell_encoding] = advanced_simulation_cell(control=control,
                                                    cell_address=excel_address,
                                                    variable_dict=variable_dict,
                                                    current_sheet_name=sheet_name)
            #case 1 
            elif type(cell_value) == float or type(cell_value) == int:
                variable_dict[sheet_cell_encoding] = np.zeros(simulation_num)+cell_value

            #case 1
            else:
                variable_dict[sheet_cell_encoding] = np.zeros(simulation_num)
                explainError(control=control, error_id="FormulaError",
                        custom_text=f"Cell {curr_cell_address}, sheet {sheet_cell_encoding} has no valid entry, default value of 0 used")
            
 
            #case 2. This is incomplete -  a formula could be wrong



    parser = Parser()

    #this adjusts the formula so it can be read and matched with our dictionary
    
    def g(x, current_sheet_name=current_sheet_name):
        """
        Note that current_sheet_name refers to the sheet_name the formula is actually on
        whereas sheet_name at this point in the process refers to the sheet the last
        item we iterated through was on.

        Hence we want to turn something like Sheet1!$A$1 to the encoding described earlier when we defined sheet_cell_encoding
        which is of the form F<ENCODING> where the F is because we need a non numeric before numbers for the Parser
        to recognise as a variable

        """
        x = x.group()
        if x[0] == "'":
            x = x.replace("'", "") #remove the ' to make it consistent with the encoding
        else: 
            pass
        if "!" in x:
            return "F" + str(int.from_bytes(x.encode(encoding="utf-8"), byteorder = sys.byteorder))
        else:
            return "F" + str(int.from_bytes((current_sheet_name+"!"+x).encode(encoding="utf-8"), byteorder = sys.byteorder))
    formula = re.sub(r"([a-zA-Z0-9]{0,}[!]{0,1}[A-Z]+[0-9]+)|(['].{1,}?['][!]{0,1}[A-Z]+[0-9]+)", g,formula)

    result = parser.parse(formula[1:]).evaluate(variable_dict)

    if first_call:
        #NEED TO LOOK INTO THIS - variable dict didn't seem to be losing values
        variable_dict.clear()
    else:
        pass
    return result
    


def advanced_simulation_cell_wrapper(control, histogram_bins=histogram_bins):
    hist_data = advanced_simulation_cell(control=control, variable_dict={}, first_call=True)

    fig, ax = plt.subplots(1, 1)

    ax.hist(hist_data, bins=histogram_bins)
    ax.grid()

    stats = scipy.stats.describe(hist_data)
    stats_names = ["runs:", "min:", "max:", "mean:", "variance:", "skewness:", "kurtosis:"]
    stats = [stats[0]] + [stats[1][0], stats[1][1]] + list(stats[2:])

    xl = xl_app()

    user_selection = xl.Selection.Address.replace("$", "")
    user_sheet_name = xl.ActiveSheet.Name

    stats_sheet_name = f"Results for cell {user_selection},  {user_sheet_name}"
    
    for sheet in xl.Sheets:
        if stats_sheet_name == sheet.Name:
            sheet.Delete()
            break
        else:
            pass
    
    xl.Sheets.Add()
    xl.ActiveSheet.Name = stats_sheet_name
    
    for i in range(len(stats)):
        xl.ActiveSheet.Range(f"A{i+1}").Value = stats_names[i]
        xl.ActiveSheet.Range(f"B{i+1}").Value = stats[i]

    return plot(fig)   

