try:
    from setuptools import setup
except ImportError:
    from distutils.core import setup

setup(
    name = 'parse_xls',
    version = '1',
    author = 'Tristan Fisher',
    author_email = 'tristan@amplify.com',
    url = 'http://bitbucket.org/syseng/python/parse_xls',
    license = 'Apache 2.0',
    install_requires = [
        'xlrd',
    ]
)