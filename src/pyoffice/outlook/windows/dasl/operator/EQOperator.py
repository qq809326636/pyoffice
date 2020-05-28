from .BaseOperator import BaseOperator

__all__ = ['EQOperator']


class EQOperator(BaseOperator):

    def __init__(self,
                 *args,
                 **kwargs):
        BaseOperator.__init__(self,
                              *args,
                              **kwargs)

        self.code = 10
        self.op = '='

