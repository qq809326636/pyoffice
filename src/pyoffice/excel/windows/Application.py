"""
Excel Application
"""

import logging

from pyoffice.decorator import singleton
from ._WinObject import _WinObject

__all__ = ['Application']


class Application(_WinObject):
    """
    Excel 应用
    """

    __instance = None

    # Field
    impl = None

    @singleton(moduleName='Application')
    def __new__(cls, *args, **kwargs):
        if cls.__instance is None:
            cls.__instance = _WinObject.__new__(cls)

            import pythoncom
            pythoncom.CoInitialize()

            if cls.impl is None:
                import win32com.client
                try:
                    cls.impl = win32com.client.GetObject(Class='Excel.Application')
                except Exception as err:
                    logging.warning(err)
                    cls.impl = win32com.client.DispatchEx('Excel.Application')
                # cls.impl.Visible = True  # default: true

        return cls.__instance

    def __init__(self):
        _WinObject.__init__(self)

    @staticmethod
    def getApplication():
        """
        获取唯一的 Excel 应用

        :return:
        """
        return Application()

    def getPid(self):
        """
        获取 Excel 应用的 PID

        :return:
        :rtype: int
        """
        import win32process
        threadId, processId = win32process.GetWindowThreadProcessId(self.impl.Hwnd)
        return processId

    def getVisible(self):
        """
        获取 Excel 显示状态

        :return:
        :rtype: bool
        """
        return self.impl.Visible

    def setVisible(self,
                   visible: bool):
        """
        设置 Excel 显示状态

        :param bool visible:
        """
        self.impl.Visible = visible

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
        from .Workbook import Workbook

        workbook = Workbook()
        workbook.open(filepath,
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
        return workbook

    def quit(self):
        """
        退出运行的 Excel 应用

        :return:
        """
        self.impl.Quit()

    def terminate(self):
        """
        终止运行的 Excel 应用

        :return:
        """
        from pyoffice.utils import ProcessUtil
        ProcessUtil.terminalProcessByPID(self.getPid())

    def getActiveWorkbook(self):
        """
        获取当前激活的工作簿

        :return: 返回一个工作簿
        :rtype: Workbook
        """
        from .Workbook import Workbook

        workbook = Workbook()
        workbook.impl = self.impl.ActiveWorkbook

        if workbook.impl is None:
            raise ValueError('The active workbook is None.')

        return workbook

    def createWorkbook(self):
        """
        创建一个工作簿

        :return: 返回一个工作簿
        :rtype: Workbook
        """
        from .Workbook import Workbook

        workbook = Workbook()
        workbook.impl = self.impl.Workbooks.Add()

        return workbook

    def getVersion(self):
        """
        获取 Excel 版本号

        :return:
        """
        return self.impl.Version

    def getExcelLimits(self):
        # if self.getVersion() == '9.0':
        #     logging.debug(f'The Excel version is 2000.')
        # elif self.getVersion() == '10.0':
        #     logging.debug(f'The Excel version is 2002/XP.')
        # elif self.getVersion() == '11.0':
        #     logging.debug(f'The Excel version is 2003.')
        # elif self.getVersion() == '12.0':
        #     logging.debug(f'The Excel version is 2007.')
        # elif self.getVersion() == '13.0':
        #     logging.debug(f'The Excel version is 2010.')
        # else:
        #     logging.debug(f'The Excel version is latest.')

        from .ExcelLimits import ExcelLimits

        limits = ExcelLimits()

        try:
            self.setVisible(True)
            wb = self.getActiveWorkbook()
        except Exception as err:
            logging.warning(err)
            wb = self.createWorkbook()

        try:
            ws = wb.getActiveWorkSheet()
        except Exception as err:
            logging.warning(err)
            ws = wb.getActiveWorkSheet()

        if self.getVersion() in ['9.0',
                                 '10.0',
                                 '11.0']:
            maxRowCount = ws.getRowByAddress(1).count()
            maxColumnCount = ws.getColumnByAddress('A').count()

            limits.maxColumnCount = maxRowCount
            limits.maxRowCount = maxColumnCount
        else:
            maxRowCount = ws.getRowByAddress(1).count()
            maxColumnCount = ws.getColumnByAddress('A').count()

            limits.maxColumnCount = maxRowCount
            limits.maxRowCount = maxColumnCount

        wb.close()

        return limits

    def getWorkbookList(self):
        from .Workbook import Workbook

        ret = list()

        for item in self.impl.Workbooks:
            wb = Workbook()
            wb.impl = item

            ret.append(wb)

        return ret

    def getWorkbookCount(self):
        return self.impl.Workbooks.Count
