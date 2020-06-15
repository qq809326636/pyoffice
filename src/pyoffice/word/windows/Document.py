from ._WinObject import *
from .Application import Application
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
              originalFormat: int = OriginalFormat.WordDocument,
              routeDocument: bool = True):
        self.impl.Close(saveChanges,
                        originalFormat,
                        routeDocument)

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
