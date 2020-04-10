"""
Workbook
"""

__all__ = ['Workbook']


class Workbook:
    def __init__(self):
        self.__app = None
        self.__workbook = None

        # init
        self.__initApplication()

    def __initApplication(self):
        if not self.__app:
            from .ExcelApplication import ExcelApplication
            self.__app = ExcelApplication()

    def setApplication(self,
                       app):
        self.__app = app

    def getApplication(self):
        return self.__app

    def display(self):
        self.__app._ExcelApplication__app.Visible = True

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
        self.__workbook = self.__app._ExcelApplication__app.Workbooks.Open(filepath,
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
