"""
Workbook
"""
from .WorkbookException import WorkbookException
from ._WinObject import _WinObject
from .constant import *

__all__ = ['Workbook']


class Workbook(_WinObject):
    """
    工作簿
    """

    def __init__(self):
        _WinObject.__init__(self)

        # init
        from .Application import Application
        self._app = Application.getApplication()
        self._attached = False

    def getApplication(self):
        """
        获取当前工作簿所属的 Excel 应用

        :return:
        :rtype: pyoffice.excel.windows.Workbook
        """
        return self._app

    def display(self):
        """
        显示当前工作簿

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
        打开一个 Excel 的工作簿

        :param str filepath: 工作簿路径
        :param bool updateLinks:
        :param bool readOnly: 是否以只读逻辑打开
        :param str format:
        :param str password: 工作簿的密码
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
        :return: 返回一个工作簿
        :rtype: Workbook
        """
        for item in self._app.impl.Workbooks:
            if item.FullName == filepath:
                self.impl = item
                self._attached = True
                break
        else:
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
        关闭当前工作簿。
        该操作不会保存关闭前所有未保存的结果。

        :return:
        """
        if self._attached:
            self.impl.Close()
        if self._app.getWorkbookCount() == 0:
            self._app.quit()

    def save(self):
        """
        保存工作簿

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
        工作簿另存为

        :param str fileName: 另存为工作簿的路径
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
        获取当前工作簿激活的工作表

        :return:
        :rtype: Worksheet
        """
        from .Worksheet import Worksheet

        workSheet = Worksheet()
        workSheet.impl = self.impl.ActiveSheet
        if workSheet.impl is None:
            raise ValueError('The active sheet is None.')
        return workSheet

    def getWorkSheetByName(self,
                           sheetName: str):
        """
        根据工作表名获取工作表

        :param str sheetName: 工作表名
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
        获取工作表数组

        :return:
        :rtype: list
        """
        from .Worksheet import Worksheet

        for item in self.impl.Worksheets:
            ws = Worksheet()
            ws.impl = item
            yield ws

    def getPath(self):
        """
        获取当前工作簿的路径

        :return: 路径
        :rtype: str
        """
        return self.impl.Path

    def isReadOnly(self):
        """
        获取当前工作表只读状态

        :return: 只读状态
        :rtype: bool
        """
        return self.impl.ReadOnly

    def getWritePassword(self):
        """
        获取当前工作簿的写入密码

        :return: 密码
        :rtype: str
        """
        return self.impl.WritePassword

    def setWritePassword(self,
                         writePassword: str):
        """
        设置写入密码

        :param str writePassword: 写入密码
        :return:
        """
        self.impl.WritePassword = writePassword

    def getAccuracyVersion(self):
        """
        获取当前精度版本

        :return:
        :rtype: int
        """
        return self.impl.AccuracyVersion

    def setAccuracyVersion(self,
                           accuracyVersion: int = AccuracyVersionEnum.LATEST):
        """
        设置当前工作簿精确版本

        :param int accuracyVersion: 版本
        :return:
        """
        self.impl.AccuracyVersionEnum = accuracyVersion

    def getActiveCell(self):
        """
        获取当前工作簿激活的单元格

        :return: 单元格
        :rtype: Cell
        """
        from .Cell import Cell

        cell = Cell()
        cell.impl = self._app.impl.ActiveCell

        return cell

    def getFirstSheet(self):
        """
        获取第一个工作表

        :return: 工作表
        :rtype: Worksheet
        """
        from .Worksheet import Worksheet

        ws = Worksheet()
        ws.impl = self.impl.Worksheets.Item(1)
        return ws

    def getLastSheet(self):
        """
        获取最后一个工作表

        :return: 工作表
        :rtype: Worksheet
        """
        from .Worksheet import Worksheet

        ws = Worksheet()
        ws.impl = self.impl.Worksheets.Item(self.getWorkSheetCount())
        return ws

    def getWorkSheetCount(self):
        """
        获取当前工作簿中工作表的数量

        :return: 工作表数量
        :rtype: int
        """
        return self.impl.Worksheets.Count
