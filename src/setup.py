from setuptools import setup, find_packages

if __name__ == '__main__':
    setup(
        name='pyoffice',
        version='1.0.0',
        description='Visualize office applications. Include in Excel, Word, etc.',
        author='LiHaibao',
        author_email='pengyou_1994@163.com',
        url='',
        download_url='',
        packages=find_packages(),
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
