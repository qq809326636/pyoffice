import importlib
import inspect
from .BaseProp import BaseProp

__all__ = ['PropFactory']


class PropFactory:

    @staticmethod
    def create(prop: str):
        moduleName = PropFactory.__module__
        module = importlib.import_module(moduleName[:moduleName.rindex('.')])

        for field in dir(module):
            cls = getattr(module, field)
            if inspect.isclass(cls) and issubclass(cls, BaseProp):
                inst = cls()
                if inst.alias.lower() == str(prop).lower() or inst.prop.lower() == str(prop).lower():
                    return inst
        else:
            raise ValueError(f'Could not found the "{prop}" property.')
