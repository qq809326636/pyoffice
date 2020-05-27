from typing import List, Optional
from .DASLCondition import DASLCondition

__all__ = ['DASLGroup']


class DASLGroup:

    def __init__(self):
        self._conditions: List[DASLGroup, DASLCondition] = []
        self._link: int = None

    def addCondition(self,
                     condition: DASLCondition):
        self._conditions.append(condition)

    def addGroup(self,
                 group: Optional['DASLGroup']):
        self._conditions.append(group)

    def removeCondition(self,
                        condition: DASLCondition):
        pass

    def removeGroup(self,
                    group: Optional['DASLGroup']):
        pass

    def _remove(self,
                item):
        pass
