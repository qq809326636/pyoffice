__all__ = ['_WinObject']


class _WinObject:

    def __init__(self):
        self._impl = None

    @property
    def impl(self):
        return self._impl

    @impl.setter
    def impl(self,
             impl):
        self._impl = impl

    @impl.deleter
    def impl(self):
        raise RuntimeError('Cannot delete "impl" property.')
