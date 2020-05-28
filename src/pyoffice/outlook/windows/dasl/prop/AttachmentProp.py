from .BaseHttpMailProp import BaseHttpMailProp

__all__ = ['AttachmentProp']


class AttachmentProp(BaseHttpMailProp):

    def __init__(self,
                 *args,
                 **kwargs):
        BaseHttpMailProp.__init__(self,
                                  *args,
                                  **kwargs)

        self.prop = 'hasattachment'
        self.alias = 'attachment'

    def format(self,
               value):
        return int(bool(value))
