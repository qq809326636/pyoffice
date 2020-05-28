from abc import ABCMeta, abstractmethod

__all__ = ['BaseOperator']


class BaseOperator(metaclass=ABCMeta):

    def __init__(self,
                 code: int = -1,
                 op: str = ''):
        self._code = code
        self._op = op

    @property
    def code(self):
        return self._code

    @code.setter
    def code(self,
             code):
        self._code = int(code)

    @property
    def op(self):
        return self._op

    @op.setter
    def op(self,
           op):
        self._op = str(op)

    def operate(self,
                prop,
                value):
        return f'{prop.getFullNamespace()} {self.op} {prop.format(value)}'
