
from setuptools import setup
from pathlib import Path


description = (
    "A tool for copying Microsoft Excel spreadsheets "
    "and conditionally deleting rows")
long_description = description
license = 'MIT'

MODULE_DIR = 'src/xlsx_copycull'


def get_constant(constant):
    setters = {
        "version": "__version__ = ",
        "author": "__author__ = ",
        "author_email": "__email__ = ",
        "url": "__website__ = "
    }
    var_setter = setters[constant]
    with open(Path(rf"{MODULE_DIR}/_constants.py"), "r") as file:
        for line in file:
            if line.startswith(var_setter):
                var = line[len(var_setter):].strip('\'\n \"')
                return var
        raise RuntimeError(f"Could not get {constant} info.")


setup(
    name='xlsx_copycull',
    version=get_constant("version"),
    packages=['xlsx_copycull'],
    package_dir={'': 'src'},
    url=get_constant("url"),
    license=license,
    author=get_constant("author"),
    author_email=get_constant("author_email"),
    description=description,
    long_description=long_description,
    install_requires=[
        'openpyxl'
    ]
)
