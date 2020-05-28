from .BaseOperator import BaseOperator

__all__ = ['GTOperator']


class GTOperator(BaseOperator):

    def __init__(self,
                 *args,
                 **kwargs):
        BaseOperator.__init__(self,
                              *args,
                              **kwargs)

        self.code = 30
        self.op = '>'

