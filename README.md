# simExcel
For simulation in excel

This uses the PyXLL excel software to build monte carlo simulation capabilities into excel. 

Note some work still needs to done to clean this up - e.g. by removing files from the PyXLL tutorials which are irrelevant to the project.


# To-do

## tkinter_frames.py and GUI
* in tkinter_frames.py, need to implement tests for input variables. E.g., if a negative variance is entered, this should raise an error message
* in tkinter_frames.py, could improve the window that appears and customize for different distributions. e.g., for normal distribution, have something where you input mean, std, and if std is negative an appropriate error message is raised
* in tkinter_frames.py, improve error messages (so it isn't just printing to the excel cell something long). (see extra features: as of 23.07.21 this is partially implemented)

## Code maintainability
* ribbon_functions.py is too long. Should split into several files based on functionality
* Need to develop unit tests

## Extra features (not sure if all of these will be helpful currently)
* Error messages. I am currently thinking having an abbreviated error message, but also having a ribbon function which you can press on it to explain the error. (Basic version implemented 23.07.2021)
* Error messages. Currently a very rudimentary return value to a cell. Perhaps better to make a pop-up tkinter box or something similar. (done for some error messages 23.07.2021. Some work needed on the UI)
* Decision variables. Implement some sort of optimisation features. (e.g. stochastic gradient descent etc)
* More complicated flow simulation
* request a feature button 
* Default value ribbon for the cell values ribbon. (done 23.07.21)
* Plotting. Have a dot on the pdf to show the mean and the median
* Default values. For the triangular distribution it is currently the mode. But this leads to silly values being displayed for some cases. E.g. for a triangular distribution with the end points at 0 and 100, with peak at 3. Displaying 3 is somewhat silly when the average will be waaaay higher than that
* fancy distributions via a drop-down menu
