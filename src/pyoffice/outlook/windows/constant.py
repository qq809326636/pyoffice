"""
constant
"""

__all__ = ['MessageType',
           'FolderType',
           'MessageCloseType']


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
