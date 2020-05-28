from .BaseHttpMailProp import BaseHttpMailProp

__all__ = ['ReceivedDateProp']


class ReceivedDateProp(BaseHttpMailProp):

    def __init__(self,
                 *args,
                 **kwargs):
        BaseHttpMailProp.__init__(self,
                                  *args,
                                  **kwargs)

        self.prop = 'created'
        self.alias = 'created'
