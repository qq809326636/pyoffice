"""
Excel Application
"""

import logging

from ._WinObject import _WinObject
from pyoffice.decorator import singleton

__all__ = ['Application']


class Application(_WinObject):
    __instance = None

    # Field
    impl = None

    @singleton(moduleName='Application')
    def __new__(cls, *args, **kwargs):
        if cls.__instance is None:
            cls.__instance = _WinObject.__new__(cls)

            if cls.impl is None:
                import win32com.client
                try:
                    cls.impl = win32com.client.GetObject(Class='Excel.Application')
                except Exception as err:
                    logging.warning(err)
                    cls.impl = win32com.client.DispatchEx('Excel.Application')
                cls.impl.Visible = True  # default: true

        return cls.__instance

    def __init__(self):
        _WinObject.__init__(self)

    @staticmethod
    def getApplication():
        return Application()

    # def __getattribute__(self, item):
    #     try:
    #         return getattr(self.impl, item)
    #     except Exception:
    #         return getattr(self, item)

    # def __getattr__(self, item):
    #     try:
    #         return getattr(self, item)
    #     except Exception:
    #         return getattr(self.impl, item)

    def getPid(self):
        """
        Get excel application process id.
        :return:
        """
        import win32process
        threadId, processId = win32process.GetWindowThreadProcessId(self.impl.Hwnd)
        return processId

    def getVisible(self):
        """
        Get the excel application visible.
        :return:
        """
        return self.impl.Visible

    def setVisible(self,
                   visible: bool):
        """
        Set the excel application visible.
        :param visible:
        :return:
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
        Open the workbook.
        :param filepath:
        :param updateLinks:
        :param readOnly:
        :param format:
        :param password:
        :param writeResPassword:
        :param ignoreReadOnlyRecommended:
        :param origin:
        :param delimiter:
        :param editable:
        :param notify:
        :param converter:
        :param addToMru:
        :param local:
        :param corruptLoad:
        :return:
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
        Quit the application.
        :return:
        """
        self.impl.Quit()

    def terminate(self):
        """
        Terminal the application
        :return:
        """
        from pyoffice.utils import ProcessUtil
        ProcessUtil.terminalProcessByPID(self.getPid())

    def getActiveWorkbook(self):
        """
        Get active workbook.
        :return:
        """
        from .Workbook import Workbook

        workbook = Workbook()
        workbook.impl = self.impl.ActiveWorkbook

        return workbook

    def createWorkbook(self):
        from .Workbook import Workbook

        workbook = Workbook()
        workbook.impl = self.impl.Workbooks.Add()

        return workbook
