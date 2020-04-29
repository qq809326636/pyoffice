from ._WinObject import *
from .constant import *

__all__ = ['Message']


class Message(_WinObject):
    def __init__(self):
        _WinObject.__init__(self)

    # For fields
    def getAttachmentCount(self):
        return self.impl.Attachments.Count

    def getAttachmentList(self):
        from .Attachment import Attachment

        for attachment in self.impl.Attachments:
            atta = Attachment()
            atta.impl = attachment
            yield atta

    def getBCC(self):
        return self.impl.BCC

    def setBCC(self,
               bcc):
        self.impl.BCC = bcc

    def getCC(self):
        return self.impl.CC

    def setCC(self,
              cc):
        self.impl.CC = cc

    def getCreationTime(self):
        return self.impl.CreationTime

    def getEntryID(self):
        return self.impl.EntryID

    def getSubject(self):
        return self.impl.Subject

    def getSender(self):
        return self.impl.Sender

    def setSender(self,
                  sender):
        from .Account import Account

        if isinstance(sender, Account):
            self.impl._oleobj_.Invoke(*(64209, 0, 8, 0, sender.impl))
        else:
            raise RuntimeError('sender must be an instance of type Account.')

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

    def save(self):
        self.impl.Save()

    def saveAs(self,
               path: str,
               saveType: int = MessageSaveType.MSGUNICODE):
        self.impl.SaveAs(path, saveType)

    def send(self):
        self.impl.Send()

    # For dependencies
    def getFolder(self):
        from .Folder import Folder

        folder = Folder()
        folder.impl = self.impl.Parent

        return folder
