from .BaseOperator import BaseOperator

__all__ = ['StartsWithOperator']


class StartsWithOperator(BaseOperator):

    def __init__(self,
                 *args,
                 **kwargs):
        BaseOperator.__init__(self,
                              *args,
                              **kwargs)

        self.code = 41
        self.op = 'ci_startswith'

