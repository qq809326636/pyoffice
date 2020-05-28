from .BaseOperator import BaseOperator

__all__ = ['LTOperator']


class LTOperator(BaseOperator):

    def __init__(self,
                 *args,
                 **kwargs):
        BaseOperator.__init__(self,
                              *args,
                              **kwargs)

        self.code = 20
        self.op = '<'

