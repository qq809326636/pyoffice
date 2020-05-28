from .BaseHttpMailProp import BaseHttpMailProp

__all__ = ['MessageProp']


class MessageProp(BaseHttpMailProp):

    def __init__(self,
                 *args,
                 **kwargs):
        BaseHttpMailProp.__init__(self,
                                  *args,
                                  **kwargs)

        self.prop = 'textdescription'
        self.alias = 'message'
