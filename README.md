# simExcel
For simulation in excel

This uses the PyXLL excel software to build monte carlo simulation capabilities into excel. 

Note some work still needs to done to clean this up - e.g. by removing files from the PyXLL tutorials which are irrelevant to the project.


# To-do

## tkinter_frames.py and GUI
* in tkinter_frames.py, need to implement tests for input variables. E.g., if a negative variance is entered, this should raise an error message
* in tkinter_frames.py, could improve the window that appears and customize for different distributions. e.g., for normal distribution, have something where you input mean, std, and if std is negative an appropriate error message is raised
* in tkinter_frames.py, improve errir messages (so it isn't just printing to the command line)

## Code maintainability
* ribbon_functions.py is too long. Should split into several files based on functionality
* Need to develop unit tests

## Extra features
* Error messages. I am currently thinking having an abbreviated error message, but also having a ribbon function which you can press on it to explain the error
* Decision variables. Implement some sort of optimisation features. (e.g. stochastic gradient descent etc)
* More complicated flow simulation
