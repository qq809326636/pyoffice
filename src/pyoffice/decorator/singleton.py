"""
Singleton
"""

import importlib
import logging

__all__ = ['singleton']

LOCK_FUNC_NAMES = ['lock',
                   'acquire']
UNLOCK_FUNC_NAMES = ['unlock',
                     'release']


def singleton(moduleName: str = '',
              className: str = '',
              lockFuncName: str = '',
              unlockFuncName: str = '',
              isMultiprocess: bool = False):
    locker = None
    try:
        if moduleName and className:
            logging.warning('Will use the specified lock.')
            module = importlib.import_module(moduleName)
            locker = getattr(module, className)()
        else:
            raise ValueError(f'The "moduleName" or "className" is empty.')
    except BaseException as err:
        logging.warning(err)

        if isMultiprocess:
            from multiprocessing import Lock
        else:
            from threading import Lock

        locker = Lock()

    # Get lock function name
    if not lockFuncName or not hasattr(locker, lockFuncName):
        for field in dir(locker):
            if not field.startswith('_') and field in LOCK_FUNC_NAMES:
                lockFuncName = field
                break
        else:
            raise RuntimeError('Must specify lock function.')

    # Get unlock function name
    if not unlockFuncName or not hasattr(locker, unlockFuncName):
        for field in dir(locker):
            if not field.startswith('_') and field in UNLOCK_FUNC_NAMES:
                unlockFuncName = field
                break
        else:
            raise RuntimeError('Must specify unlock function.')

    def synchronized(func):

        func.__locker__ = locker

        def wrapper(*args,
                    **kwargs):
            try:
                getattr(func.__locker__, lockFuncName)()
                return func(*args,
                            **kwargs)
            except Exception as err:
                logging.error(err)
            finally:
                getattr(func.__locker__, unlockFuncName)()

        return wrapper

    return synchronized
