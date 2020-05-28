from .BaseHttpMailProp import BaseHttpMailProp

__all__ = ['RecipientProp']


class RecipientProp(BaseHttpMailProp):

    def __init__(self,
                 *args,
                 **kwargs):
        BaseHttpMailProp.__init__(self,
                                  *args,
                                  **kwargs)

        self.prop = 'recipient'
        self.alias = 'recipient'
