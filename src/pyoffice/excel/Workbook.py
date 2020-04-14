import platform

if platform.system().lower() == 'windows':
    from .windows import Workbook
else:
    raise RuntimeError(f'This {platform.system()} platform does not support.')
