from .BaseHttpMailProp import BaseHttpMailProp

__all__ = ['CCProp']


class CCProp(BaseHttpMailProp):

    def __init__(self,
                 *args,
                 **kwargs):
        BaseHttpMailProp.__init__(self,
                                  *args,
                                  **kwargs)

        self.prop = 'displaycc'
        self.alias = 'cc'
