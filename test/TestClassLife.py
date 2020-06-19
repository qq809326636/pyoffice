import pytest
import inspect


class Clazz:
    def __new__(cls, *args, **kwargs):
        print('__new__')
        return object.__new__(cls)

    def __init__(self):
        print('__init__')

    def __del__(self):
        print('__del__')

    def __call__(self, *args, **kwargs):
        print('__call__')

    def __getattr__(self, item):
        print('__getattr__')


class CustomerMethodDict(dict):

    def __setitem__(self, key, value):
        print('=' * 80)
        print('CustomerMethodDict.__setitem__')
        print(f'key: {key}')
        print(f'value: {value}')

        super().__setitem__(key, value)


class CustomerMeta(type):

    def __init__(self,
                 name,
                 bases,
                 dict):
        print('=' * 80)
        print('CustomerMeta.__init__')
        print(f'name: {name}')
        print(f'bases: {bases}')
        print(f'dict: {dict}')

        type.__init__(self,
                      name,
                      bases,
                      dict)

    def __del__(self):
        print('=' * 80)
        print('CustomerMeta.__del__')

    def __new__(mcs,
                name,
                bases,
                dict):
        print('=' * 80)
        print('CustomerMeta.__new__')
        print(f'name: {name}')
        print(f'bases: {bases}')
        print(f'dict: {dict}')

        return type.__new__(mcs,
                            name,
                            bases,
                            dict)

    @classmethod
    def __prepare__(metacls, name, bases):
        print('=' * 80)
        print('CustomerMeta.__prepare__')
        print(f'name: {name}')
        print(f'bases: {bases}')
        return CustomerMethodDict({
            'impl': None
        })

    def __call__(self, *args, **kwargs):
        print('=' * 80)
        print('CustomerMeta.__call__')
        print(f'args: {args}')
        print(f'kwargs: {kwargs}')

        return super().__call__(*args,
                                **kwargs)


class CustomerClass(metaclass=CustomerMeta):
    enum = None
    enum1 = 111
    enum2 = str

    def __init__(self):
        print('CustomerClass.__init__')

        self._x = None
        self.__x = None
        self.x = None

    def echo(self):
        print('CustomerClass.echo')

    def echo(self,
             val: str):
        print(val)


# class SubCustomerClass(CustomerClass):
#
#     def __init__(self):
#         CustomerClass.__init__(self)
#
#     def echo(self):
#         print('SubCustomerClass.echo')


# class TestClassLife:
#
#     def test_clazz(self):
#         cls = Clazz()
#         cls()
#         x = cls.x
#
#     def test_customermeta(self):
#         inst = CustomerClass()
#         inst.echo()
#         print(inst.impl)
#         print(inst.x)


print('=' * 80)
inst = CustomerClass()
# inst.echo()
# print(inst.impl)
# print(inst.x)
