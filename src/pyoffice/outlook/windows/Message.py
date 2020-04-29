from ._WinObject import *
from .constant import MessageCloseType

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

    def getSender(self):
        return self.impl.Sender

    def setSender(self,
                  sender):
        self.impl._oleobj_.Invoke(*(64209, 0, 8, 0, sender))

    # For methods
    def close(self,
              closeType: int = MessageCloseType.SAVE):
        self.impl.Close(closeType)

    def copy(self):
        message = Message()
        message.impl = self.impl.Copy()
        return message

    def delete(self):
        self.impl.Delete()

    def display(self,
                model: bool = False):
        self.impl.Display(model)

    def forward(self):
        message = Message()
        message.impl = self.impl.Forward()
        return message

    def move(self,
             folder):
        self.impl.Move(folder.impl)

    def reply(self):
        message = Message()
        message.impl = self.impl.Reply()
        return message

    def replyAll(self):
        message = Message()
        message.impl = self.impl.ReplyAll()
        return message

    # For dependencies
    def getFolder(self):
        from .Folder import Folder

        folder = Folder()
        folder.impl = self.impl.Parent

        return folder
