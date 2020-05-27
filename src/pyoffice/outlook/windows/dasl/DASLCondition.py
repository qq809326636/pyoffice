from .DASLUtil import DASLUtil
from .constant import *
from .DASLDate import DASLDate

__all__ = ['DASLCondition']


class DASLCondition:

    def __init__(self):
        self._prop = None
        self._op: int = 10
        self._val = None
        self._link: int = -1

    @property
    def prop(self):
        return self._prop

    @prop.setter
    def prop(self,
             prop):
        self._prop = str(prop)

    @property
    def op(self):
        return self._op

    @op.setter
    def op(self,
           op):
        self._op = int(op)

    @property
    def val(self):
        return self._val

    @val.setter
    def val(self,
            right):
        self._val = right

    @property
    def link(self):
        return self._link

    @link.setter
    def link(self,
             link):
        self._link = int(link)

    def toPartOfDASL(self):
        ret = ''
        ns, val = DASLUtil.getPropertyByKey(self._prop, self._val)
        if self._op == 40:
            ret = '"{}" {} \'%{}%\''.format(
                ns,
                DASLOperatorEnum.LIKE,
                val
            )
        elif self._op == 41:
            ret = '"{}" {} \'{}%\''.format(
                ns,
                DASLOperatorEnum.LIKE,
                val
            )
        elif self._op == 42:
            ret = '"{}" {} \'%{}\''.format(
                ns,
                DASLOperatorEnum.LIKE,
                val
            )
        else:
            op = DASLUtil.getOperatorByCode(self._op)

            if isinstance(val, (str, DASLDate)):
                val = f'\'{str(val)}\''

            ret = '"{}" {} {}'.format(
                ns,
                op,
                val
            )

        link = DASLUtil.getLinkByCode(self._link)
        return f' {link} ( {ret} )'

    def __str__(self):
        return self.toPartOfDASL()

    def __repr__(self):
        return self.toPartOfDASL()
