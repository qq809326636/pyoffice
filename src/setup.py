import os
import setuptools
from distutils.core import setup
from distutils.extension import Extension
from Cython.Build import cythonize
import platform

# module_list = [
#     Extension(name='winexcel',
#               sources=[
#                   r'pyoffice\excel\windows\_WinObject.py',
#                   r'pyoffice\excel\windows\Application.py',
#                   r'pyoffice\excel\windows\Cell.py',
#                   r'pyoffice\excel\windows\Column.py',
#                   r'pyoffice\excel\windows\PivotTable.py',
#                   r'pyoffice\excel\windows\Range.py',
#                   r'pyoffice\excel\windows\Row.py',
#                   r'pyoffice\excel\windows\Table.py',
#                   r'pyoffice\excel\windows\Workbook.py',
#                   r'pyoffice\excel\windows\Worksheet.py'
#               ]),
#     Extension(name='winprocutils',
#               sources=[
#                   r'pyoffice\utils\processmenager\windows\ProcessUtil.py'
#               ])
# ]

module_list = list()
excludeFileList = ['setup.py']
for dirpath, dirnames, filenames in os.walk('.'):
    for filename in filenames:
        if filename.endswith('.py') and not filename.startswith('_') and filename not in excludeFileList:
            module_list.append(
                Extension(os.path.splitext(filename)[0],
                          [os.path.join(dirpath, filename)])
            )

setup(
    name='pyoffice',
    version='1.0.0',
    description='Visualize office applications. Include in Excel, Word, etc.',
    author='LiHaibao',
    author_email='pengyou_1994@163.com',
    url='',
    download_url='',
    ext_modules=cythonize(module_list=module_list,
                          compiler_directives={
                              'language_level': 3
                          }),
    classifiers=[
        'Environment :: Win32 (MS Windows)',
        'Development Status :: 3 - Alpha',
        'Topic :: Software Development :: Libraries'
    ],
    license='GPL',
    keywords=['office',
              'excel',
              'outlook',
              'word'],
    platforms=['Windows']
)
