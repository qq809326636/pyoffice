from .BaseOperator import BaseOperator

__all__ = ['PhraseMatchOperator']


class PhraseMatchOperator(BaseOperator):

    def __init__(self,
                 *args,
                 **kwargs):
        BaseOperator.__init__(self,
                              *args,
                              **kwargs)

        self.code = 42
        self.op = 'ci_phrasematch'


