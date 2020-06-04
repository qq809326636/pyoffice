from .constant import *

__all__ = ['Util']


class Util:

    @staticmethod
    def columnLableToIndex(lable: str):
        index = 0
        for item in lable.upper():
            if 'A' <= item <= 'Z':
                index = index * 26 + ord(item) - ord('A') + 1
            else:
                raise ValueError(f'The column "{lable}" lable is wrong.')

        return index

    @staticmethod
    def columnLableFromIndex(index: int):
        if index > SheetMax.MAX_COL:
            raise IndexError(f'The column index has exceeded the maximum number of columns.')
        if index < 1:
            raise IndexError('Column index must start from 1.')

        labelList = list()
        while index > 0:
            index -= 1
            lab = index % 26
            labelList.insert(0, lab)
            index = index // 26

        return ''.join([chr(i + 65) for i in labelList])
