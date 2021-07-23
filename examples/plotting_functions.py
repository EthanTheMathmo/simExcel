from pyxll import xl_func, plot
import matplotlib.pyplot as plt
import numpy as np


distributions_dictionary = {"N": 1} #for mapping inputs to 

@xl_func
def simple_plot():
    # Data for plotting
    t = np.arange(0.0, 2.0, 0.01)
    s = 1 + np.sin(2 * np.pi * t)

    # Create the figure and plot the data
    fig, ax = plt.subplots()
    ax.plot(t, s)

    ax.set(xlabel='time (s)', ylabel='voltage (mV)',
           title='About as simple as it gets, folks')
    ax.grid()

    # Display the figure in Excel
    plot(fig)

