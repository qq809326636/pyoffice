from .BaseOperator import BaseOperator

__all__ = ['LEOperator']


class LEOperator(BaseOperator):

    def __init__(self,
                 *args,
                 **kwargs):
        BaseOperator.__init__(self,
                              *args,
                              **kwargs)

        self.code = 21
        self.op = '<='


