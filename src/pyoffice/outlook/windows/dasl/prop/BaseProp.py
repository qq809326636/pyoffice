from abc import abstractmethod

__all__ = ['BaseProp']


class BaseProp:

    def __init__(self,
                 prop: str = '',
                 namespace: str = ''):
        self._prop: str = prop
        self._namespace: str = namespace
        self._alias: str = prop

    @property
    def prop(self):
        return self._prop

    @prop.setter
    def prop(self,
             prop):
        self._prop = str(prop)

    @property
    def namespace(self):
        return self._namespace

    @namespace.setter
    def namespace(self,
                  namespace):
        self._namespace = str(namespace)

    @property
    def alias(self):
        return self._alias

    @alias.setter
    def alias(self,
              alias):
        self._alias = alias

    @abstractmethod
    def format(self,
               value):
        raise RuntimeError('Must to implement this method.')

    def getFullNamespace(self):
        fullNamespace = ':'.join([self.namespace, self.prop])
        return f'"{fullNamespace}"'
