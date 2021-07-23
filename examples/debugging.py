from pyxll import xl_menu
import ptvsd
 
@xl_menu("Attach To VS Code")
def vscode_enable_attach():
    # Allow VS Code to attach to Python using the default port number 5678
    ptvsd.enable_attach()
 
    # If your port number is not 5678 you can specify it as follows:
    #ptvsd.enable_attach(address=('localhost', YOUR_PORT_NUMBER))
 
    # Pause the program until a remote debugger is attached
    # (This is optional)
    ptvsd.wait_for_attach()