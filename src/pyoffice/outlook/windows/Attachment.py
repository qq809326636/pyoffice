from ._WinObject import _WinObject

__all__ = ['Attachment']


class Attachment(_WinObject):

    def __init__(self):
        _WinObject.__init__(self)

    # For fields
    def getDisplayName(self):
        return self.impl.DisplayName

    def getFileName(self):
        return self.impl.FileName

    def getIndex(self):
        return self.impl.Index

    def getPathName(self):
        return self.impl.PathName

    def getSize(self):
        return self.impl.Size

    def getType(self):
        return self.impl.Type

    # For methods
    def delete(self):
        self.impl.Delete()

    def getTemporaryFilePath(self):
        return self.impl.GetTemporaryFilePath()

    def saveAsFile(self,
                   path: str):
        self.impl.SaveAsFile(path)

    # For dependencies
