__all__ = ['_WinObject']


class _WinObject:

    def __init__(self):
        self._impl = None
        self._parent = None

    @property
    def impl(self):
        return self._impl

    @impl.setter
    def impl(self,
             impl):
        if not self._impl:
            self._impl = impl
        else:
            raise RuntimeError('The "impl" is not empty and cannot be reassigned.')

    @impl.deleter
    def impl(self):
        raise RuntimeError('Cannot delete "impl" property.')

    @property
    def parent(self):
        return self._parent

    @parent.setter
    def parent(self,
               parent):
        if not self._parent:
            self._parent = parent
        else:
            raise RuntimeError('The "parent" is not empty and cannot be reassigned.')

    @parent.deleter
    def parent(self):
        raise RuntimeError('Cannot delete "parent" property.')
