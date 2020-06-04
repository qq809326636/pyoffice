"""
Table
"""
from ._WinObject import _WinObject

__all__ = ['Table']


class Table(_WinObject):
    """
    数据表
    """

    def __init__(self):
        _WinObject.__init__(self)

    def getName(self):
        return self.impl.DisplayName
