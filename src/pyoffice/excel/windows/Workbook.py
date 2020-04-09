"""
Workbook
"""

__all__ = ['Workbook']


class Workbook:
    def __init__(self):
        self.__workbook = None

    @property
    def workbook(self):
        return self.__workbook

    @workbook.setter
    def workbook(self,
                 workbook):
        self.__workbook = workbook
