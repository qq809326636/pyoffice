"""
Builder
"""

from .Expression import Expression
from .Group import Group

__all__ = ['Builder']


class Builder:

    @staticmethod
    def build(expr: dict):
        dasl = Builder._build(expr)
        return dasl

    @staticmethod
    def _build(expr: dict):
        if 'group' in expr.keys():
            # if is a group
            groupExpr = expr['group']
            group = Group()
            group.linker = groupExpr['linker']
            group.setLeft(Builder._build(groupExpr['left']))
            group.setRight(Builder._build(groupExpr['right']))
            return group
        else:
            dasl = Expression()
            dasl.prop = expr['prop']
            dasl.op = expr['op']
            dasl.value = expr['value']
            return dasl
