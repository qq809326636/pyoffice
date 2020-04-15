"""
Table
"""

from ._WinObject import _WinObject

__all__ = ['PivotTable']


class PivotTable(_WinObject):

    def __init__(self):
        _WinObject.__init__(self)
