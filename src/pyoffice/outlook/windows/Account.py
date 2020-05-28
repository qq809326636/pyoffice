from ._WinObject import *
from .constant import FolderType

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

    def createMessage(self):
        from .Message import Message
        from .constant import MessageType

        msg = Message()
        msg.impl = self._session.Application.CreateItem(MessageType.MAILITEM)

        return msg

    def getDefaultFolder(self):
        from .Folder import Folder

        folder = Folder()
        folder.impl = self._folder
        return folder

    def createFolder(self,
                     folderName: str,
                     folderType: int = FolderType.FOLDER_INBOX):
        from .Folder import Folder
        from .FolderUtil import FolderUtil

        rootFolder = Folder()
        rootFolder.impl = self._folder
        if not FolderUtil.hasFolderExists(rootFolder,
                                          folderName):
            folder = Folder()
            folder.impl = self._folder.Folders.Add(folderName,
                                                   folderType)
            return folder
        else:
            raise RuntimeError(f'The "{folderName}" folder already exists.')
