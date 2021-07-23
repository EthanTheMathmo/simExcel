from pyxll import xl_func


@xl_func
def hello(name):
    return f"Hello {name}"