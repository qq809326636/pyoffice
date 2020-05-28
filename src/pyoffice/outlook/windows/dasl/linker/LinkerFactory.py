import importlib
import inspect
from .BaseLinker import BaseLinker

__all__ = ['LinkerFactory']


class LinkerFactory:

    @staticmethod
    def create(linker):
        moduleName = LinkerFactory.__module__
        module = importlib.import_module(moduleName[:moduleName.rindex('.')])

        for field in dir(module):
            cls = getattr(module, field)
            if inspect.isclass(cls) and issubclass(cls, BaseLinker):
                inst = cls()
                if inst.code == linker or inst.op.lower() == str(linker).lower():
                    return inst
        else:
            raise RuntimeError('Could not found the operator.')
