# Copyright (c) 2010-2025 openpyxl
import weakref


class Singleton(type):
    """
    Singleton metaclass
    Based on Python Cookbook 3rd Edition Recipe 9.13
    Only one instance of a class can exist. Does not work with __slots__
    """

    def __init__(cls, *args, **kw):
        super().__init__(*args, **kw)
        cls.__instance = None

    def __call__(cls, *args, **kw):
        if cls.__instance is None:
            cls.__instance = super().__call__(*args, **kw)
        return cls.__instance


class Cached(type):
    """
    Caching metaclass
    Child classes will only create new instances of themselves if
    one doesn't already exist. Does not work with __slots__
    """

    def __init__(cls, *args, **kw):
        super().__init__(*args, **kw)
        cls.__cache = weakref.WeakValueDictionary()

    def __call__(cls, *args):
        if args in cls.__cache:
            return cls.__cache[args]
        obj = super().__call__(*args)
        cls.__cache[args] = obj
        return obj
