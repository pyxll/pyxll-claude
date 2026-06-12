"""
Excel functions exposed to PyXLL.

Write functions decorated with @xl_func here, then click **Load Functions**
in the Claude AI task pane to register them with Excel immediately.

Example:

    from pyxll import xl_func

    @xl_func
    def greet(name: str) -> str:
        '''Return a greeting string.'''
        return f"Hello, {name}!"
"""
from pyxll import xl_func
