from .BaseHttpMailProp import BaseHttpMailProp

__all__ = ['SenderProp']


class SenderProp(BaseHttpMailProp):

    def __init__(self,
                 *args,
                 **kwargs):
        BaseHttpMailProp.__init__(self,
                                  *args,
                                  **kwargs)

        self.prop = 'fromname'
        self.alias = 'sender'
