# simExcel
For simulation in excel. Implements recursive algorithm to parse mathematical formuli and perform markov chain monte carlo simulation across an Excel workbook. Currently works for elementary formuli built from +/-/exp/log/sin/cos/etc. Built as a demo replacement to the oracle crystal ball software. 

This uses the PyXLL excel software to build monte carlo simulation capabilities into excel. 

UPDATE: this is now deprecated, as we switched to using the opensource alternatives (xlwings and openpyxl)

# To-do


## tkinter_frames.py and GUI
* need to improve appearance of pop ups

## Advanced simulation
### Preprocessing
* Page names now all work with int encoding!! :) as of 28.07.21
* Empty cells cause problems. (e.g., if I have SUM(A1:B5) but B4 has no entries. 

### extra functionality
* things like AVG, SUM, etc TO-DO this is obviously useful for spreadsheets




## General performance
* see https://stackoverflow.com/questions/10714251/how-to-avoid-using-select-in-excel-vba/10718179#10718179 I think there are some places where I am using .Active, .Selection etc where I don't strictly need to (I think this is now sorted as of 28.07.21)

## Code maintainability
* ribbon_functions.py is too long. Should split into several files based on functionality
* Need to develop unit tests
* There are some places where I am switching the activesheet instead of using x.Worksheet(sheet_name).<...> which is more inefficient and is only like that because I didn't know about .Worksheets
* the error messages pop up system could be simplified as a class system. (e.g., most basic being the PopupWindow, and then making the specific error messages being extensions of that

## input distributions
### error catching input
* lots of things to work on here if there's time. 
* Need to personalise the error catching. Example, for beta_prime distirbution, we need beta>1 for the mean to be defined, so if we pass in something with beta <= 1, the default value breaks
* Catching errors in input, e.g. if they input non numerical characters (perhaps regex it so we get out the right number of floats?)

### remove/add distributions
* need a button to allow removal of distributions





## Extra features (not sure if all of these will be helpful currently)
* comprehensive error catching. currently it only catches bits and bobs and there aren't tests designed for each function
* Decision variables. Implement some sort of optimisation features. (e.g. stochastic gradient descent etc)
* request a feature button 
* Plotting. Have a dot on the pdf to show the mean and the median


## Extra features, completed modulo work making it neater
* More complicated flow simulation (compound formulas across multiple cells done 24.07.21. Probably should implement some common functions like SUM and AVG but for various reasons this would be a little fiddly) (done, but without functions like SUM and AVG)
* Default value ribbon for the cell values ribbon. (done 23.07.21)
* Default values. For the triangular distribution it is currently the mode. But this leads to silly values being displayed for some cases. E.g. for a triangular distribution with the end points at 0 and 100, with peak at 3. Displaying 3 is somewhat silly when the average will be waaaay higher than that (done)
* fancy distributions via a drop-down menu (done)
* error messages tinkering and invalid selection. Some bugs need to be fixed (e.g., entering one entry 0 for normal distribution raises an error, but entering 0, doesn't give an error and it should. 
* Decision for error messages and handling. Instead of raising error messages can just have an empty return so it does nothing. The tkinter frames give more info but are a bit awkward (which becomes apparent when using them). - maybe need to experiment with the other options
