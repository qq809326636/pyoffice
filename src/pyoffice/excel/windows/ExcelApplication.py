"""
Excel Application
"""
__all__ = ['ExcelApplication']


class ExcelApplication:

    def __init__(self):
        import win32com.client

        self.__app = win32com.client.Dispatch('Excel.Application')

    def getVisible(self):
        """
        Get the excel application visible.
        :return:
        """
        return self.__app.Visible

    def setVisible(self,
                   visible: bool):
        """
        Set the excel application visible.
        :param visible:
        :return:
        """
        self.__app.Visible = visible

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
        self.__app.Quit()
