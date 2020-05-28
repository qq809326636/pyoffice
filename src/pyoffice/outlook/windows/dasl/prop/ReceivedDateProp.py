from .BaseHttpMailProp import BaseHttpMailProp
import datetime
import time

__all__ = ['SentDateProp']


class SentDateProp(BaseHttpMailProp):

    def __init__(self,
                 *args,
                 **kwargs):
        BaseHttpMailProp.__init__(self,
                                  *args,
                                  **kwargs)

        self.prop = 'date'
        self.alias = 'sentdate'

    def format(self,
               value):
        fmt = '%Y-%m-%d %H:%M:%S'
        if isinstance(value, datetime.datetime):
            value = value.strftime(fmt)
        elif isinstance(value, time.struct_time):
            value = time.strftime(fmt, value)
        return super().format(value)
