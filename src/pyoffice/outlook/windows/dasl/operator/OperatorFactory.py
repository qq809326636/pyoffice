import importlib
import inspect
from .BaseOperator import BaseOperator

__all__ = ['OperatorFactory']


class OperatorFactory:

    @staticmethod
    def create(op):
        moduleName = OperatorFactory.__module__
        module = importlib.import_module(moduleName[:moduleName.rindex('.')])

        for field in dir(module):
            cls = getattr(module, field)
            if inspect.isclass(cls) and issubclass(cls, BaseOperator):
                inst = cls()
                if inst.code == op or inst.op.lower() == str(op).lower():
                    return inst
        else:
            raise RuntimeError('Could not found the operator.')
