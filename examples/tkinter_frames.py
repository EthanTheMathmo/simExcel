import tkinter as tk
from pyxll import xl_app

#helpful variables
from meta_variables import distributions_dictionary, id_location, screen_freeze_disabled


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

        #where we'll store the result
        self.form_result = "0,1" #default mean 0, s.d. 1

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


        self.xl.ScreenUpdating = self.screen_freeze_disabled #this ensures no screen flickering from switching the active sheet

        self.xl.Worksheets(self.distrInfoPageName).Activate()

        #set the relevant values on the distrInfoSheet
        self.xl.ActiveSheet.Range(self.user_selection).Value = self.entry_value.get() + "," + self.distribution_id

        # "".join([self.form_result["Mean"],self.form_result["Standard deviation"], "N"])
        #return the active sheet to the user's original page
        self.xl.Worksheets(self.userCurrentPageName).Activate()
        
        #set the user's selected cells to have a numerical value equal to the mean
       #set the user's selected cells to have a numerical value equal to the mean
        params = self.entry_value.get().split(",")
        self.check_params(param_array=params)
        if self.distribution_id == "N":
            self.xl.ActiveSheet.Range(self.user_selection).Value = float(self.entry_value.get().split(",")[0])
        elif self.distribution_id == "T":
            self.xl.ActiveSheet.Range(self.user_selection).Value = float(self.entry_value.get().split(",")[0])*float(self.entry_value.get().split(",")[2])
        elif self.distribution_id == "E":
            self.xl.ActiveSheet.Range(self.user_selection).Value = float(params[0])+float(params[1]) 
        else:
            self.xl.ActiveSheet.Range(self.user_selection).Value = "<Need to add default value for this distribution. Search for this error in tkinter_frames.py>"

        self.button1.config(relief=tk.SUNKEN)
        self.button1.after(150, lambda: self.button1.config(relief=tk.RAISED))
        

        self.xl.ScreenUpdating = True #so that the screen updates once this operation is performed

        #shuts the tk window
        self.master.destroy()
    
    def check_params(self, param_array):
        if len(param_array) != self.distributions_dictionary[self.distribution_id]["num_params"]:
            self.xl.ActiveSheet.Range(self.user_selection).Value = "wrong number of parameters"
            self.master.destroy()
        else:
            pass

        return



# class NormalData(tk.Frame):

#     def __init__(self, master, control, id_location, distribution_id="N"):
#         super().__init__(master)
#         self.control=control
#         self.xl = xl_app()
#         #the current selection of the user
#         self.user_selection = self.xl.Selection.Range

#         #where the name of the sheet containing distribution input is
#         self.id_location = id_location

#         #where we'll store the result
#         self.form_result = "0,1" #default mean 0, s.d. 1

#         #name of the page where we'll store the distribution info
#         self.distrInfoPageName = self.xl.ActiveSheet.Range(self.id_location).Value

#         #name of user's current page
#         self.userCurrentPageName = self.xl.ActiveSheet.Name

#         #distribution id
#         self.distribution_id = distribution_id

#         self.initUI()

#     def initUI(self):
#         # allow the widget to take the full space of the root window
#         self.pack(fill=tk.BOTH, expand=True)

#         # Create a tk.Entry control and place it using the 'pack' method
#         #adds the mean of the normal distribution
#         #add a button underneath for submitting
#         self.entry_value = tk.StringVar()
#         self.entry = tk.Entry(self, textvar=self.entry_value)
#         self.entry.pack()
#         self.button1 = tk.Button(self, text="Input mean and standard deviation", command=self.on_button1)
#         self.button1.pack()


#         # Allow the first column in the grid to stretch horizontally
#         self.columnconfigure(0, weight=1)

#     def updateDistrSheet(self):
#         #change the active sheet to the distribution info sheet for the user's page
#         self.xl.Worksheets(self.distrInfoPageName).Activate()

#         #set the relevant values on the distrInfoSheet
#         self.xl.ActiveSheet.Range(self.user_selection).Value = 1
#         # "".join([self.form_result["Mean"],self.form_result["Standard deviation"], "N"])
#         #return the active sheet to the user's original page
#         self.xl.Worksheets(self.userCurrentPageName).Activate()

#         self.quit()



#     def on_button1(self):
#         self.value = self.mean_entry.get()

#         #below is code so we look like we pressed the button
#         self.button1.config(relief=tk.SUNKEN)
#         self.button1.after(150, lambda: self.button1.config(relief=tk.RAISED))
#         self.updateDistrSheet()


