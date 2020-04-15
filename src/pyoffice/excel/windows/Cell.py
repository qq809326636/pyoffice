"""
Cell
"""

from ._WinObject import _WinObject

__all__ = ['Cell']


class Cell(_WinObject):

    def __init__(self):
        _WinObject.__init__(self)
