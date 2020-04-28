from ._WinObject import *

__all__ = ['Message']


class Message(_WinObject):
    def __init__(self):
        _WinObject.__init__(self)

    # For fields
    def getBCC(self):
        return self.impl.BCC

    def getCC(self):
        return self.impl.CC

    def getEntryID(self):
        return self.impl.EntryID

    def getSubject(self):
        return self.impl.Subject
