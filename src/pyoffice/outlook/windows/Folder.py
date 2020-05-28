from ._WinObject import *

__all__ = ['Folder']


class Folder(_WinObject):

    def __init__(self):
        _WinObject.__init__(self)

    # For fields
    def getAddressBookName(self):
        return self.impl.AddressBookName

    def getCurrentView(self):
        return self.impl.CurrentView

    def getEntryID(self):
        return self.impl.EntryID

    def getFolderPath(self):
        return self.impl.FolderPath

    def getFolderCount(self):
        return self.impl.Folders.Count

    def getFolderList(self):
        for item in self.impl.Folders:
            folder = Folder()
            folder.impl = item
            yield folder

    def getFolderNameList(self):
        for folder in self.getFolderList():
            yield folder.getName()

    def getFolderByName(self, name: str):
        folder = Folder()
        folder.impl = self.impl.Folders(name)
        return folder

    def getMessageCount(self):
        return self.impl.Items.Count

    def getMessageList(self):
        from .Message import Message

        for item in self.impl.Items:
            msg = Message()
            msg.impl = item
            yield msg

    def getName(self):
        return self.impl.Name

    def getUnReadItemCount(self):
        return self.impl.UnReadItemCount

    # For methods
    def copyTo(self, folder):
        self.impl.CopyTo(folder.impl)

    def delete(self):
        self.impl.Delete()

    def display(self):
        self.impl.Display()

    def getCustomIcon(self):
        return self.impl.GetCustomIcon()

    def moveTo(self, folder):
        self.impl.MoveTo(folder.impl)

    def restrict(self,
                 filter):
        pass

    def setCustomIcon(self, picture):
        self.impl.SetCustomIcon(picture)

    def query(self,
              ql: dict):
        from .dasl import Builder, DASLPrefix
        from .Message import Message

        query = Builder.build(ql)
        query = f'{DASLPrefix.PREFIX}{query}'
        ret = self.impl.Items.Restrict(query)
        for item in ret:
            msg = Message()
            msg.impl = item
            yield msg
