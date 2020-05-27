from .constant import *

__all__ = ['DASLUtil']


class DASLUtil:

    @staticmethod
    def getLinkByCode(code: int):
        code = int(code)

        for k, v in LinkEnum.items():
            if code == v['code']:
                return v['op']
        raise ValueError(f'Could not found the code "{code}"')

    @staticmethod
    def getOperatorByCode(code: int):
        code = int(code)

        for k, v in OperatorEnum.items():
            if code == v['code']:
                return v['op']
        raise ValueError(f'Could not found the code "{code}"')

    @staticmethod
    def getPropertyByKey(key: str,
                         val=None):
        if key not in PropertyEnum.keys():
            raise ValueError(f'Could not found the key "{key}"')

        return ':'.join([
            'urn',
            'schemas',
            PropertyEnum[key]['parent'],
            PropertyEnum[key]['ref']
        ]), PropertyEnum[key]['type'](val)
