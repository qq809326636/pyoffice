from .BaseCalendarProp import BaseCalendarProp

__all__ = ['BCCProp']


class BCCProp(BaseCalendarProp):

    def __init__(self,
                 *args,
                 **kwargs):
        BaseCalendarProp.__init__(self,
                                  *args,
                                  **kwargs)

        self.prop = 'resources'
        self.alias = 'bcc'
