from .operator import *
from .prop import *

__all__ = ['Expression']


class Expression:

    def __init__(self,
                 prop: (str, int) = '',
                 op: (str, int) = '',
                 value=''):
        self._prop: str = prop
        self._op: (str, int) = op
        self._value = value

    @property
    def prop(self):
        if isinstance(self._prop, BaseProp):
            return self._prop
        return PropFactory.create(self._prop)

    @prop.setter
    def prop(self,
             prop):
        self._prop = prop

    @property
    def op(self):
        if isinstance(self._op, BaseOperator):
            return self._op
        return OperatorFactory.create(self._op)

    @op.setter
    def op(self,
           op):
        self._op = op

    @property
    def value(self):
        return self._value

    @value.setter
    def value(self,
              value):
        self._value = value

    def toString(self):
        return self.op.operate(self.prop,
                               self.value)

    def __str__(self):
        return self.toString()

    def __repr__(self):
        return self.toString()
