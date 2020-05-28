from .BaseHttpMailProp import BaseHttpMailProp
from .constant import *

__all__ = ['ImportanceProp']


class ImportanceProp(BaseHttpMailProp):

    def __init__(self,
                 *args,
                 **kwargs):
        BaseHttpMailProp.__init__(self,
                                  *args,
                                  **kwargs)

        self.prop = 'importance'
        self.alias = 'importance'

    def format(self,
               value):
        value = int(value)

        if value in ImportanceEnum.getValues():
            return value
        else:
            raise ValueError(f'The value is not in {[i for i in ImportanceEnum.getValues()]}')
