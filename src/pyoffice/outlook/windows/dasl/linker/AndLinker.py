from .BaseLinker import BaseLinker

__all__ = ['AndLinker']


class AndLinker(BaseLinker):
    def __init__(self,
                 *args,
                 **kwargs):
        BaseLinker.__init__( self,
                             *args,
                            *kwargs)

        self.code = 10
        self.op = 'AND'

    def link(self,
             left,
             right=None):
        return f'( {str(left)}) {self.op} ( {str(right)} )'
