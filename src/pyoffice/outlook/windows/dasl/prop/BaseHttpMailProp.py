from .BaseProp import BaseProp

__all__ = ['BaseHttpMailProp']


class BaseHttpMailProp(BaseProp):

    def __init__(self,
                 *args,
                 **kwargs):
        BaseProp.__init__(self,
                          *args,
                          **kwargs)

        self.namespace = 'urn:schemas:httpmail'

    def format(self,
               value):
        return f'\'{str(value)}\''
