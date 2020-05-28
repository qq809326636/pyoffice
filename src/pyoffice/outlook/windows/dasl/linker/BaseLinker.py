__all__ = ['BaseLinker']


class BaseLinker:

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

    def link(self,
             left,
             right=None):
        if right is not None:
            raise RuntimeError('The root linker must be')

        return str(left)
