from .BaseOperator import BaseOperator

__all__ = ['GEOperator']


class GEOperator(BaseOperator):

    def __init__(self,
                 *args,
                 **kwargs):
        BaseOperator.__init__(self,
                              *args,
                              **kwargs)

        self.code = 31
        self.op = '>='
