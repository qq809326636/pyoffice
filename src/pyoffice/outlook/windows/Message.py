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

    def getContent(self):
        return self.impl.Body

    def setContent(self,
                   content: str):
        self.impl.Body = content

    def getContentFormat(self):
        return self.impl.BodyFormat

    def setContentFormat(self,
                         contentFormat):
        self.impl.BodyFormat = contentFormat

    def getCreationTime(self):
        return self.impl.CreationTime

    def getCategories(self):
        return self.impl.Categories

    def setCategories(self,
                      categories: str):
        self.impl.Categories = categories

    def getDownloadState(self):
        return self.impl.MessageDownloadState

    def getEntryID(self):
        return self.impl.EntryID

    def getExpiryTime(self):
        return self.impl.ExpiryTime

    def getHTMLContent(self):
        return self.impl.HTMLBody

    def setHTMLContent(self,
                       html: str):
        self.impl.HTMLBody = html

    def getImportance(self):
        return self.impl.Importance

    def setImportance(self,
                      importance: int = MessageImportance.NORMAL):
        self.impl.Importance = importance

    def getLastModificationTime(self):
        return self.impl.LastModificationTime

    def getMarkForDownload(self):
        return self.impl.MarkForDownload

    def setMarkForDownloads(self,
                            status: int = MessageRemoteStatus.UNMARKED):
        self.impl.MarkForDownload = status

    def getReceivedTime(self):
        return self.impl.ReceivedTime

    def getReminderTime(self):
        return self.impl.ReminderTime

    def getRemoteStatus(self):
        return self.impl.RemoteStatus

    def setRemoteStatus(self,
                        status: int = MessageRemoteStatus.UNMARKED):
        self.impl.RemoteStatus = status

    def getRTFContent(self):
        return self.impl.RTFBody

    def setRTFContent(self,
                      content: str):
        self.impl.RTFBody = content

    def hasSaved(self):
        return self.impl.Saved

    def getSubject(self):
        return self.impl.Subject

    def getSender(self):
        sender = self.impl.Sender
        address = sender.Address
        if address:
            return address
        else:
            return str(sender)

    def setSender(self,
                  sender):
        from .Account import Account

        if isinstance(sender, Account):
            self.impl._oleobj_.Invoke(*(64209, 0, 8, 0, sender.impl))
        else:
            raise RuntimeError('sender must be an instance of type Account.')

    def getSendUsingAccount(self):
        from .Account import Account

        acc = Account()
        acc.impl = self.impl.SendUsingAccount
        return acc

    def setSendUsingAccount(self,
                            account):
        self.impl.SendUsingAccount = account.impl

    def getSize(self):
        return self.impl.Size

    def getTo(self):
        return self.impl.To

    def setTo(self,
              to):
        self.impl.To = to

    def getReadStatus(self):
        return not self.impl.UnRead

    def setReadStatus(self,
                      status):
        self.impl.UnRead = not status

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
