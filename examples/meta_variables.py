"""
This is for variables we use across several scripts
"""
import scipy.stats

distributions_dictionary = {"N":{"num_params":2, "scipy_handle":scipy.stats.norm, "params":"mean, variance","Name": "Normal Distribution"},
                            "C":{"num_params":2, "scipy_handle":scipy.stats.cauchy, "params": "mean, scaling","Name": "Cauchy"},
                            "T":{"num_params":3, "scipy_handle":scipy.stats.triang, "params": "c, loc, scale", "Name": "Triangular"},
                            "E":{"num_params":2, "scipy_handle":scipy.stats.expon, "params": "loc, scale", "Name":"Exponential"}}


id_location = "$A$1" #note the value of the id_location will
#at some point need to be changed to a hidden location
screen_freeze_disabled = True #for debugging, screen freezing often causes problems
#set to false to freeze screen while function operations are carried out

#number of simulations performed by default
simulation_num = 15000

#dictionary matching error codes to what the error is
PNumEr_str = "This error means you entered the wrong number of parameters for the distribution selected"
error_messages_dictionary = {"PNumEr":PNumEr_str}