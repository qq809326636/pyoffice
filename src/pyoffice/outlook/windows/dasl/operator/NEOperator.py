from .BaseOperator import BaseOperator

__all__ = ['NEOperator']


class NEOperator(BaseOperator):

    def __init__(self,
                 *args,
                 **kwargs):
        BaseOperator.__init__(self,
                              *args,
                              **kwargs)

        self.code = 11
        self.op = '<>'


