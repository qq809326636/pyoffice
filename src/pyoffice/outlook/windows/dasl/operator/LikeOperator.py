from .BaseOperator import BaseOperator

__all__ = ['LikeOperator']


class LikeOperator(BaseOperator):

    def __init__(self,
                 *args,
                 **kwargs):
        BaseOperator.__init__(self,
                              *args,
                              **kwargs)

        self.code = 40
        self.op = 'like'
