from .Application import Application
from .Bookmark import *
from .Range import *
from ._WinObject import *
from .constant import *

__all__ = ['Document']


class Document(_WinObject):

    def __init__(self):
        _WinObject.__init__(self)

        self._app = Application.getApplication()
        self.__filepath = ''

    def setFilepath(self,
                    filepath: str):
        self.__filepath = filepath

    def getFilepath(self) -> str:
        return self.__filepath

    def create(self,
               template=None,
               newTemplate=None,
               documentType=None,
               visible: bool = True):
        doc = Document()
        param = dict()

        if template is not None:
            param.update({
                'Template': template
            })

        if newTemplate is not None:
            param.update({
                'NewTemplate': newTemplate
            })

        if documentType is not None:
            param.update({
                'DocumentType': documentType
            })

        if visible is not None:
            param.update({
                'Visible': visible
            })

        self._app.impl.Documents.Add(**param)

    def open(self,
             *,
             confirmConversions: bool = True,
             readOnly: bool = False,
             addToRecentFiles: bool = True,
             passwordDocument: str = '',
             passwordTemplate: str = '',
             revert: bool = False,
             writePasswordDocument: str = '',
             writePasswordTemplate: str = '',
             format: int = OpenFormat.OpenFormatAuto,
             encoding: int = MsoEncoding.EncodingUTF8,
             visible: bool = True,
             # openConflictDocument: bool = False,
             openAndRepair: bool = False,
             documentDirection: int = DocumentDirection.LeftToRight,
             noEncodingDialog: bool = False):
        if visible:
            self._app.setVisible(visible)
        param = {
            'FileName': self.__filepath,
            'ConfirmConversions': confirmConversions,
            'ReadOnly': readOnly,
            'AddToRecentFiles': addToRecentFiles,
            'PasswordDocument': passwordDocument,
            'PasswordTemplate': passwordTemplate,
            'Revert': revert,
            'WritePasswordDocument': writePasswordDocument,
            'WritePasswordTemplate': writePasswordTemplate,
            'Format': format,
            'Encoding': encoding,
            'Visible': visible,
            # 'OpenConflictDocument': openConflictDocument,
            'OpenAndRepair': openAndRepair,
            'DocumentDirection': documentDirection,
            'NoEncodingDialog': noEncodingDialog
        }
        self.impl = self._app.impl.Documents.Open(**param)

    def close(self,
              saveChanges: int = SaveOptions.DoNotSaveChanges,
              originalFormat: int = OriginalFormat.OriginalDocumentFormat,
              routeDocument: bool = True):
        self.impl.Close(saveChanges,
                        originalFormat,
                        routeDocument)

    def closePrintPreview(self):
        self.impl.ClosePrintPreview()

    def save(self):
        self.impl.Save()

    def saveAs(self,
               filepath: str,
               fileFormat: int = SaveFormat.FormatDocument,
               lookComments: bool = False,
               password: str = '',
               addToRecentFiles: bool = True,
               writePassword: str = '',
               readOnlyRecommended: bool = False,
               embedTrueTypeFonts: bool = True,
               saveNativePictureFormat: bool = True,
               saveFormsData: bool = True,
               saveAsAOCELetter: bool = True,
               encoding: int = MsoEncoding.EncodingUTF8,
               insertLineBreaks: bool = True,
               allowSubstitutions: bool = False,
               lineEnding: int = LineEndingType.CRLF,
               addBiDiMarks: bool = True,
               compatibilityMode: int = CompatibilityMode.Current):

        param = {
            'FileName': filepath,
            'FileFormat': fileFormat,
            'LockComments': lookComments,
            'Password': password,
            'AddToRecentFiles': addToRecentFiles,
            'WritePassword': writePassword,
            'ReadOnlyRecommended': readOnlyRecommended,
            'EmbedTrueTypeFonts': embedTrueTypeFonts,
            'SaveNativePictureFormat': saveNativePictureFormat,
            'SaveFormsData': saveFormsData,
            'SaveAsAOCELetter': saveAsAOCELetter,
            'Encoding': encoding,
            'InsertLineBreaks': insertLineBreaks,
            'AllowSubstitutions': allowSubstitutions,
            'LineEnding': lineEnding,
            'AddBiDiMarks': addBiDiMarks,
            'CompatibilityMode': compatibilityMode,
        }

        self.impl.SaveAs2(**param)

    def active(self):
        self.impl.Activate()

    def getCreator(self) -> int:
        return self.impl.Creator

    def getName(self):
        return self.impl.Name

    def getFullName(self) -> str:

        return self.impl.FullName

    def getPath(self) -> str:
        return self.impl.Path

    def hasPassword(self) -> bool:
        return self.impl.HasPassword

    def isReadOnly(self) -> bool:
        return self.impl.ReadOnly

    def isSaved(self) -> bool:
        return self.impl.Saved

    def getWords(self) -> list:
        return list(self.impl.Words)

    def getRange(self,
                 *,
                 start: int = None,
                 end: int = None) -> Range:
        param = dict()

        if start is not None:
            param.update({
                'Start': start
            })

        if end is not None:
            param.update({
                'End': end
            })

        rg = Range()
        rg.impl = self.impl.Range(**param)

        return rg

    def autoFormat(self):
        self.impl.AutoFormat()

    def goto(self,
             *,
             what: int = None,
             which: int = None,
             count: int = None,
             name: str = None) -> Range:
        param = dict()

        if what is not None:
            param.update({
                'What': what
            })

        if which is not None:
            param.update({
                'Which': which
            })

        if count is not None:
            param.update({
                'Count': count
            })

        if name is not None:
            param.update({
                'Name': name
            })

        rg = Range()
        rg.impl = self.impl.GoTo(**param)

        return rg

    def printPreview(self):
        self.impl.PrintPreview()

    def printOut(self,
                 *,
                 background: bool = False,
                 append: bool = False,
                 rg: Range = None,
                 outputFileName: bool = False,
                 frompages: int = None,
                 topages: int = None,
                 item=None,
                 copies: int = None,
                 pages: str = '',
                 pageType: int = None,
                 printToFile: bool = False,
                 collate: bool = False,
                 filename: str = '',
                 activePrinterMacGX=None,
                 manualDuplexPrint: bool = False,
                 printZoomColumn: int = None,
                 printZoomRow: int = None,
                 printZoomPaperWidth: int = None,
                 printZoomPaperHeight: int = None):
        param = {
            'Background': background,
            'Append': append,
            'Range': rg,
            'OutputFileName': outputFileName,
            'From': frompages,
            'To': topages,
            'Item': item,
            'Copies': copies,
            'Pages': pages,
            'PageType': pageType,
            'PrintToFile': printToFile,
            'Collate': collate,
            'FileName': filename,
            'ActivePrinterMacGX': activePrinterMacGX,
            'ManualDuplexPrint': manualDuplexPrint,
            'PrintZoomColumn': printZoomColumn,
            'PrintZoomRow': printZoomRow,
            'PrintZoomPaperWidth': printZoomPaperWidth,
            'PrintZoomPaperHeight': printZoomPaperHeight
        }

        self.impl.PrintOut(**param)

    def protect(self,
                *,
                t: int = None,
                noReset: bool = False,
                password: str = '',
                useIRM=None,
                enforceStyleLock=None,
                ):
        param = {
            'Type': t,
            'NoReset': noReset,
            'Password': password,
            'UseIRM': useIRM,
            'EnforceStyleLock': enforceStyleLock
        }

        self.impl.Protect(**param)

    def unprotect(self,
                  password: str):
        self.impl.Unprotect(password)

    def redo(self,
             times: int = 1):
        self.impl.Redo(times)

    def undo(self,
             times: int = 1):
        self.impl.Undo(times)

    def undoClear(self):
        self.impl.UndoClear()

    def select(self):
        self.impl.Select()

    def getBookmarkList(self) -> list:
        for item in self.impl.Bookmarks:
            bookmark = Bookmark()
            bookmark.impl = item
            yield bookmark

    def getContent(self) -> Range:
        rg = Range()
        rg.impl = self.impl.Content
        return rg

    def getSectionList(self) -> list:
        from .Section import Section

        for item in self.impl.Sections:
            sec = Section()
            sec.impl = item
            yield sec

    def getParagraphList(self) -> list:
        from .Paragraph import Paragraph

        for item in self.impl.Paragraphs:
            par = Paragraph()
            par.impl = item
            yield item
