from ._WinObject import *

__all__ = ['Account']


class Account(_WinObject):

    def __init__(self):
        _WinObject.__init__(self)

        # fields
        self._session = None
        self._folder = None

    @_WinObject.impl.setter
    def impl(self, impl):
        super(Account, Account).impl.__set__(self, impl)
        self._session = self.impl.Session
        self._folder = self._session.Folders(self.impl.DisplayName)

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

    # For dependencies
    def getFolderCount(self):
        return self._folder.Folders.Count

    def getFolderList(self):
        from .Folder import Folder

        for item in self._folder.Folders:
            folder = Folder()
            folder.impl = item
            yield folder

    def getFolderByName(self,
                        name):
        from .Folder import Folder

        folder = Folder()
        folder.impl = self._folder.Folders(name)
        return folder
