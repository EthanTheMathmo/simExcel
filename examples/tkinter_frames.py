"""
Implements tkinter frames


Relevant info for avoiding circular dependencies:
imports from meta_variables and error_functions

is imported by ribbon_functions.py
"""


import tkinter as tk
from pyxll import xl_app
from error_functions import default_values


#helpful variables
from meta_variables import distributions_dictionary, id_location, screen_freeze_disabled
from meta_variables import error_messages_dictionary



class DistributionData(tk.Frame):

    def __init__(self, master, control, id_location=id_location, screen_freeze_disabled=screen_freeze_disabled, 
                    distribution_id="N",distributions_dictionary=distributions_dictionary):
        super().__init__(master)
        self.initUI()

        self.xl = xl_app()
        #the current selection of the user
        self.user_selection = self.xl.Selection.Address

        #where the name of the sheet containing distribution input is
        self.id_location = id_location


        #name of the page where we'll store the distribution info
        self.distrInfoPageName = self.xl.ActiveSheet.Range(self.id_location).Value

        #name of user's current page
        self.userCurrentPageName = self.xl.ActiveSheet.Name

        #distribution id
        self.distribution_id = distribution_id

        #determines whether we freeze the screen (for debugging easier if we don't)
        self.screen_freeze_disabled = screen_freeze_disabled

        #for distribution information
        self.distributions_dictionary = distributions_dictionary

        #control variable
        self.control = control

    def initUI(self):
        # allow the widget to take the full space of the root window
        self.pack(fill=tk.BOTH, expand=True)

        # Create a tk.Entry control and place it using the 'grid' method
        self.entry_value = tk.StringVar()
        self.entry = tk.Entry(self, textvar=self.entry_value)
        self.entry.grid(column=0, row=0, padx=10, pady=10, sticky="ew")
        self.entry.pack()
        self.button1 = tk.Button(self, text="Input distribution information", command=self.on_button1)
        self.button1.pack()

        # Allow the first column in the grid to stretch horizontally
        self.columnconfigure(0, weight=1)

    def on_button1(self, *args):
        """Called when the tk.Entry's text is changed"""

        #this checks if the input distribution info has the right number of parameters
        #should update check_params to test info such as if std > 0 etc
        params = self.entry_value.get().split(",")
        if self.check_params(param_array=params):
            self.master.destroy() #ends the process if the wrong number of parameters are entered
        else:
            self.xl.ScreenUpdating = self.screen_freeze_disabled #this ensures no screen flickering from switching the active sheet

            self.xl.Worksheets(self.distrInfoPageName).Activate()

            #set the relevant values on the distrInfoSheet
            self.xl.ActiveSheet.Range(self.user_selection).Value = self.entry_value.get() + "," + self.distribution_id

            # "".join([self.form_result["Mean"],self.form_result["Standard deviation"], "N"])
            #return the active sheet to the user's original page
            self.xl.Worksheets(self.userCurrentPageName).Activate()
            
            #set the user's selected cells to have a numerical value equal to the mean
        #set the user's selected cells to have a numerical value equal to the mean
            default_values(control=self.control, selection=self.user_selection,
                        distribution_id=self.distribution_id,
                        params = params)

            self.button1.config(relief=tk.SUNKEN)
            self.button1.after(150, lambda: self.button1.config(relief=tk.RAISED))
            

            self.xl.ScreenUpdating = True #so that the screen updates once this operation is performed

            #shuts the tk window
            self.master.destroy()
    
    def check_params(self, param_array):
        if len(param_array) != self.distributions_dictionary[self.distribution_id]["num_params"]:
            self.xl.ActiveSheet.Range(self.user_selection).Value = "PNumEr"
            return True
        else:
            """
            TO-DO. Implement tests for inputs (e.g.: standard deviation should be greater 
            than zero)
            """
            pass

        return


