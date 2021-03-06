"""
constant
"""

__all__ = ['FolderType',
           'MessageType',
           'MessageCloseType',
           'MessageSaveType',
           'MessageDownloadState',
           'MessageImportance',
           'MessageRemoteStatus',
           'OutlookNamespaces']


class MessageType:
    APPOINTMENTITEM = 1  # An AppointmentItem object.
    CONTACTITEM = 2  # A ContactItem object.
    DISTRIBUTIONLISTITEM = 7  # A DistListItem object.
    JOURNALITEM = 4  # A JournalItem object.
    MAILITEM = 0  # A MailItem object.
    NOTEITEM = 5  # A NoteItem object.
    POSTITEM = 6  # A PostItem object.
    TASKITEM = 3  # A TaskItem object.


class FolderType:
    FOLDER_CALENDAR = 9  # The Calendar folder.
    FOLDER_CONFLICTS = 19  # The Conflicts folder (subfolder of the Sync Issues folder). Only available for an Exchange account.
    FOLDER_CONTACTS = 10  # The Contacts folder.
    FOLDER_DELETEDITEMS = 3  # The Deleted Items folder.
    FOLDER_DRAFTS = 16  # The Drafts folder.
    FOLDER_INBOX = 6  # The Inbox folder.
    FOLDER_JOURNAL = 11  # The Journal folder.
    FOLDER_JUNK = 23  # The Junk E-Mail folder.
    FOLDER_LOCALFAILURES = 21  # The Local Failures folder (subfolder of the Sync Issues folder). Only available for an Exchange account.
    FOLDER_MANAGEDEMAIL = 29  # The top-level folder in the Managed Folders group. For more information on Managed Folders, see the Help in Microsoft Outlook. Only available for an Exchange account.
    FOLDER_NOTES = 12  # The Notes folder.
    FOLDER_OUTBOX = 4  # The Outbox folder.
    FOLDER_SENTMAIL = 5  # The Sent Mail folder.
    FOLDER_SERVERFAILURES = 22  # The Server Failures folder (subfolder of the Sync Issues folder). Only available for an Exchange account.
    FOLDER_SUGGESTEDCONTACTS = 30  # The Suggested Contacts folder.
    FOLDER_SYNCISSUES = 20  # The Sync Issues folder. Only available for an Exchange account.
    FOLDER_TASKS = 13  # The Tasks folder.
    FOLDER_TODO = 28  # The To Do folder.
    PUBLICFOLDER_SALLPUBLICFOLDERS = 18  # The All Public Folders folder in the Exchange Public Folders store. Only available for an Exchange account.
    FOLDER_RSSFEEDS = 25  # The RSS Feeds folder.


class MessageCloseType:
    DISCARD = 1  # Changes to the document are discarded.
    PROMPTFORSAVE = 2  # User is prompted to save documents.
    SAVE = 0  # Documents are saved.


class MessageSaveType:
    DOC = 4  # Microsoft Office Word format (.doc)
    HTML = 5  # HTML format (.html)
    ICAL = 8  # iCal format (.ics)
    MHTML = 10  # MIME HTML format (.mht)
    MSG = 3  # Outlook message format (.msg)
    MSGUNICODE = 9  # Outlook Unicode message format (.msg)
    RTF = 1  # Rich Text format (.rtf)
    TEMPLATE = 2  # Microsoft Outlook template (.oft)
    TXT = 0  # Text format (.txt)
    VCAL = 7  # VCal format (.vcs)
    VCARD = 6  # VCard format (.vcf)


class MessageDownloadState:
    FULL = 1  # Full item has been downloaded.
    HEADER_ONLY = 0  # Only the header has been downloaded.


class MessageImportance:
    HIGH = 2  # Item is marked as high importance.
    LOW = 0  # Item is marked as low importance.
    NORMAL = 1  # Item is marked as medium importance.

    @staticmethod
    def getDesc(status: int):
        for filed in dir(MessageImportance):
            if not filed.startswith('_'):
                val = getattr(MessageImportance, filed, None)
                if val and val == status:
                    return filed
        else:
            raise RuntimeError(f'The "{status}" is wrong.Please check it.')


class MessageRemoteStatus:
    MARKEDFORCOPY = 3  # Item is marked to be copied.
    MARKEDFORDELETE = 4  # Item is marked for deletion.
    MARKEDFORDOWNLOAD = 2  # Item is marked for download.
    REMOTESTATUSNONE = 0  # No remote status has been set.
    UNMARKED = 1  # Item is not marked.


class OutlookNamespaces:
    MAPI_PROPTAG='https://schemas.microsoft.com/mapi/proptag'  # Outlook item objects, AddressEntry, AddressList, Attachment, ExchangeDistributionList, ExchangeUser, Folder, Recipient, and Store objects.
    MAPI_ID='https://schemas.microsoft.com/mapi/id'  # (Same as above)
    MAPI_STRING='https://schemas.microsoft.com/mapi/string'  # (Same as above)
    EXCHANGE='https://schemas.microsoft.com/exchange'  # (Same as above)
    URN_OFFICE='urn:schemas-microsoft-com:office:office'  # Outlook item objects
    URN_OUTLOOK='urn:schemas-microsoft-com:office:outlook'  # Outlook item objects
    DAV='DAV:'  # Outlook item objects
    URN_CALENDAR='urn:schemas:calendar'  # Outlook item objects
    URN_CONTACTS='urn:schemas:contacts'  # Outlook item objects
    URN_HTTP_MAIL='urn:schemas:httpmail'  # Outlook item objects
    URN_MAIL_HEADER='urn:schemas:mailheader'  # Outlook item objects
