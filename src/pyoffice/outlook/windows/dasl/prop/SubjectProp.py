from .BaseHttpMailProp import BaseHttpMailProp

__all__ = ['SubjectProp']


class SubjectProp(BaseHttpMailProp):

    def __init__(self,
                 *args,
                 **kwargs):
        BaseHttpMailProp.__init__(self,
                                  *args,
                                  **kwargs)

        self.prop = 'subject'
        self.alias = 'subject'
