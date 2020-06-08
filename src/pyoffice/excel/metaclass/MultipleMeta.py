from .MultiDict import *

__all__ = ['MultipleMeta']


class MultipleMeta(type):
    """
    Metaclass that allows multiple dispatch of methods
    """

    def __new__(cls,
                clsname,
                bases,
                clsdict):
        return type.__new__(cls, clsname, bases, dict(clsdict))

    @classmethod
    def __prepare__(cls,
                    clsname,
                    bases):
        return MultiDict()
