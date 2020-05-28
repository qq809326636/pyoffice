from .BaseProp import BaseProp

__all__ = ['BaseCalendarProp']


class BaseCalendarProp(BaseProp):

    def __init__(self,
                 *args,
                 **kwargs):
        BaseProp.__init__(self,
                          *args,
                          **kwargs)

        self.namespace = 'urn:schemas:calendar'

    def format(self,
               value):
        return str(value)
