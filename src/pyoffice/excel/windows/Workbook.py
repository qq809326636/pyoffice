"""
Workbook
"""
from .WorkbookException import WorkbookException
from ._WinObject import _WinObject
from .constant import *

__all__ = ['Workbook']


class Workbook(_WinObject):
    """
    Workbook
    """

    def __init__(self):
        _WinObject.__init__(self)

        # init
        from .Application import Application
        self._app = Application.getApplication()

    def getApplication(self):
        """
        Get application.

        :return:
        :rtype: Application
        """
        return self._app

    def display(self):
        """
        Show workbook

        :return:
        """
        self._app.setVisible(True)

    def open(self,
             filepath: str,
             updateLinks: bool = False,
             readOnly: bool = False,
             format=None,
             password: str = '',
             writeResPassword=None,
             ignoreReadOnlyRecommended=False,
             origin=None,
             delimiter=None,
             editable: bool = True,
             notify: bool = False,
             converter=None,
             addToMru=None,
             local=None,
             corruptLoad=None):
        """
        Open the workbook.

        :param str filepath:
        :param bool updateLinks:
        :param bool readOnly:
        :param str format:
        :param str password:
        :param str writeResPassword:
        :param bool ignoreReadOnlyRecommended:
        :param origin:
        :param delimiter:
        :param bool editable:
        :param bool notify:
        :param converter:
        :param addToMru:
        :param local:
        :param corruptLoad:
        :return:
        :rtype: Workbook
        """
        wb = self._app.impl.Workbooks.Open(filepath,
                                           updateLinks,
                                           readOnly,
                                           format,
                                           password,
                                           writeResPassword,
                                           ignoreReadOnlyRecommended,
                                           origin,
                                           delimiter,
                                           editable,
                                           notify,
                                           converter,
                                           addToMru,
                                           local,
                                           corruptLoad)
        self.impl = wb

    def close(self):
        """
        Close this workbook without save.

        :return:
        """
        self.impl.Close()

    def save(self):
        """
        Save the workbook

        :return:
        """
        self.impl.Save()

    def saveAs(self,
               fileName: str,
               fileFormat: int = XLFileFormatEnum.xlOpenXMLWorkbook,
               password: str = '',
               writeResPassword: str = '',
               readOnlyRecommended: bool = True,
               createBackup: bool = True,
               accessMode: int = XlSaveAsAccessMode.xlShared,
               conflictResolution: int = XlSaveConflictResolution.xlLocalSessionChanges,
               addToMru: bool = False,
               textCodepage=None,
               textVisualLayout=None,
               local: bool = True):
        """
        The workbook save as other document.

        :param str fileName:
        :param int fileFormat:
        :param str password:
        :param str writeResPassword:
        :param bool readOnlyRecommended:
        :param bool createBackup:
        :param int accessMode:
        :param int conflictResolution:
        :param bool addToMru:
        :param textCodepage:
        :param textVisualLayout:
        :param bool local:
        :return:
        """
        self.impl.SaveAs(fileName,
                         fileFormat,
                         password,
                         writeResPassword,
                         readOnlyRecommended,
                         createBackup,
                         accessMode,
                         conflictResolution,
                         addToMru,
                         textCodepage,
                         textVisualLayout,
                         local)

    def getActiveWorkSheet(self):
        """
        Get Active WorkSheet

        :return:
        :rtype: Worksheet
        """
        from .Worksheet import Worksheet

        workSheet = Worksheet()
        workSheet.impl = self.impl.ActiveSheet
        return workSheet

    def getWorkSheetByName(self,
                           sheetName: str):
        """
        Get WorkSheet By Name

        :param sheetName:
        :return:
        """
        from .Worksheet import Worksheet

        ws = Worksheet()
        for item in self.impl.Worksheets:
            if sheetName == item.Name:
                ws.impl = item
                return ws
        else:
            raise WorkbookException(f'No worksheet with this name {sheetName} found.')

    def getWorkSheetList(self) -> list:
        """
        Get WorkSheet List

        :return:
        """
        from .Worksheet import Worksheet

        for item in self.impl.Worksheets:
            ws = Worksheet()
            ws.impl = item
            yield ws

    def getPath(self):
        """
        Get file path

        :return:
        """
        return self.impl.Path

    def isReadOnly(self):
        """
        Get workbook read only attribute.

        :return:
        """
        return self.impl.ReadOnly

    def getWritePassword(self):
        """
        Get workbook password.

        :return:
        """
        return self.impl.WritePassword

    def setWritePassword(self,
                         writePassword: str):
        """
        Set workbook password

        :param writePassword:
        :return:
        """
        self.impl.WritePassword = writePassword

    def getAccuracyVersion(self):
        """
        Get accuracy version.

        :return:
        """
        return self.impl.AccuracyVersion

    def setAccuracyVersion(self,
                           accuracyVersion: int = AccuracyVersionEnum.LATEST):
        """
        Set accuracy version.

        :param accuracyVersion:
        :return:
        """
        self.impl.AccuracyVersionEnum = accuracyVersion

    def getActiveCell(self):
        """
        Get active cell.

        :return:
        """
        from .Cell import Cell

        cell = Cell()
        cell.impl = self._app.impl.ActiveCell

        return cell

    def getFirstSheet(self):
        """
        Get first sheet.

        :return:
        """
        from .Worksheet import Worksheet

        ws = Worksheet()
        ws.impl = self.impl.Worksheets.Item(1)
        return ws

    def getLastSheet(self):
        """
        Get last sheet.

        :return:
        """
        from .Worksheet import Worksheet

        ws = Worksheet()
        ws.impl = self.impl.Worksheets.Item(self.getWorkSheetCount())
        return ws

    def getWorkSheetCount(self):
        """
        Get worksheet count

        :return:
        """
        return self.impl.Worksheets.Count


