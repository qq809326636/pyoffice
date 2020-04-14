"""
Excel Application
"""

from ._WinObject import _WinObject

__all__ = ['ExcelApplication']


class ExcelApplication(_WinObject):

    def __init__(self):
        _WinObject.__init__(self)

        import win32com.client

        self._impl = win32com.client.Dispatch('Excel.Application')

    def getPid(self):
        import win32process
        threadId, processId = win32process.GetWindowThreadProcessId(self._impl.Hwnd)
        return processId

    def getVisible(self):
        """
        Get the excel application visible.
        :return:
        """
        return self._impl.Visible

    def setVisible(self,
                   visible: bool):
        """
        Set the excel application visible.
        :param visible:
        :return:
        """
        self._impl.Visible = visible

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
        workbook.setApplication(self)
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
        self._impl.Quit()
