from ._WinObject import *

__all__ = ['Explorer']


class Explorer(_WinObject):

    def __init__(self):
        _WinObject.__init__(self)

    def display(self):
        self.impl.Display()
