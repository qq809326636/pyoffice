from ._WinObject import *

__all__ = ['Account']


class Account(_WinObject):

    def __init__(self):
        _WinObject.__init__(self)

    # For Fields
    def getAccountType(self):
        return self.impl.AccountType

    def getClass(self):
        return self.impl.Class

    def getCurrentUser(self):
        return self.impl.CurrentUser

    def getDisplayName(self):
        return self.impl.DisplayName

    def getSmtpAddress(self):
        return self.impl.SmtpAddress

    def getUserName(self):
        return self.impl.UserName

    # For Methods

