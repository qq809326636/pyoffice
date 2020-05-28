from .BaseHttpMailProp import BaseHttpMailProp

__all__ = ['ReadProp']


class ReadProp(BaseHttpMailProp):

    def __init__(self,
                 *args,
                 **kwargs):
        BaseHttpMailProp.__init__(self,
                                  *args,
                                  **kwargs)

        self.prop = 'read'
        self.alias = 'read'

    def format(self,
               value):
        return int(bool(value))
