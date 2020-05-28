from .BaseLinker import BaseLinker

__all__ = ['OrLinker']


class OrLinker(BaseLinker):
    def __init__(self,
                 *args,
                 **kwargs):
        BaseLinker.__init__(self,
                            *args,
                            *kwargs)

        self.code = 11
        self.op = 'OR'

    def link(self,
             left,
             right=None):
        return f'( {str(left)} ) {self.op} ( {str(right)} )'
