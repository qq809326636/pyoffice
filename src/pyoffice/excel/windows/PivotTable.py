"""
Table
"""

from ._WinObject import _WinObject

__all__ = ['PivotTable']


class PivotTable(_WinObject):
    """
    透视表
    """

    def __init__(self):
        _WinObject.__init__(self)
