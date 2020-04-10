"""
WorkSheet
"""

__all__ = ['WorkSheet']


class WorkSheet:

    def __init__(self):
        self.__workSheet = None

    def getName(self):
        return self.__workSheet.Name
