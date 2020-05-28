__all__ = ['ImportanceEnum']


class BaseEnum:

    @classmethod
    def getKeys(cls):
        for field in dir(cls):
            if not field.startswith('_') and field not in ['getKeys',
                                                           'getValues']:
                yield field

    @classmethod
    def getValues(cls):
        for field in dir(cls):
            if not field.startswith('_') and field not in ['getKeys',
                                                           'getValues']:
                yield getattr(cls, field)


class ImportanceEnum(BaseEnum):
    LOW = 0
    NORMAL = 1
    HIGH = 2
