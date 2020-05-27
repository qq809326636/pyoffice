import logging
from .DASLCondition import DASLCondition
from .DASLGroup import DASLGroup
from .constant import *

__all__ = ['DASLBuilder']


class DASLBuilder:

    def __init__(self):
        self._conditions = list()

    def add(self,
            item):
        self._conditions.append(item)

    def remove(self,
               item):
        return self._conditions.remove(item)

    def build(self,
              isAdvancedSearch: bool = False):
        # for item in self._conditions:
        #     logging.debug(f'item value is "{item}"')
        #
        #     #
        #     subConditions = list()
        #     if isinstance(item, DASLCondition):
        #         # if condition
        #         pass
        #     elif isinstance(item, DASLGroup):
        #         # if group
        #         pass

        return DASLPrefix.PREFIX + ''.join([str(item) for item in self._conditions])
