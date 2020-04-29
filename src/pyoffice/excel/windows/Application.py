"""
Excel Application
"""

import logging

from ._WinObject import _WinObject

__all__ = ['Application']


class Application(_WinObject):

    def __init__(self):
        _WinObject.__init__(self)

        import win32com.client
        try:
            self.impl = win32com.client.GetActiveObject(Class='Excel.Application')
        except Exception as err:
            logging.warning(err)
            self.impl = win32com.client.DispatchEx('Excel.Application')
        self.impl.Visible = True  # default: true

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
        workbook.parent = self
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
