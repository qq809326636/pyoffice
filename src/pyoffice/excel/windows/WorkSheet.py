"""
WorkSheet
"""

__all__ = ['WorkSheet']


class WorkSheet:

    def __init__(self):
        self.__workSheet = None

    def getName(self):
        """
        Get worksheet name
        :return:
        """
        return self.__workSheet.Name
