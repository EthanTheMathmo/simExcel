# simExcel
For simulation in excel

This uses the PyXLL excel software to build monte carlo simulation capabilities into excel. 

Note some work still needs to done to clean this up - e.g. by removing files from the PyXLL tutorials which are irrelevant to the project.


# To-do

## tkinter_frames.py and GUI
* in tkinter_frames.py, need to implement tests for input variables. E.g., if a negative variance is entered, this should raise an error message
* in tkinter_frames.py, could improve the window that appears and customize for different distributions. e.g., for normal distribution, have something where you input mean, std, and if std is negative an appropriate error message is raised
* in tkinter_frames.py, improve error messages (so it isn't just printing to the excel cell something long). (see extra features: as of 23.07.21 this is partially implemented)

## Advanced simulation
This basically works as of 27.07.21, with some minor things to work on if there's time:
- some page names break it (e.g., a page name with brackets, or * or + because these confuse the formula reading)
- ideally have a function which runs through the cells and checks in advance if the formula is valid (i.e., to avoid the user having to face nasty error messages)
- Some work has been done on error management, e.g., if a cell has no entries in it or in the distributions sheet, an error message is raised


## Code maintainability
* ribbon_functions.py is too long. Should split into several files based on functionality
* Need to develop unit tests
* There are some places where I am switching the activesheet instead of using x.Worksheet(sheet_name).<...> which is more inefficient and is only like that because I didn't know about .Worksheets
* the error messages pop up system could be simplified as a class system. (e.g., most basic being the PopupWindow, and then making the specific error messages being extensions of that

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
