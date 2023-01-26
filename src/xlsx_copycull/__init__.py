# Copyright (c) 2021-2022, James P. Imes. All rights reserved.

"""
Tools for generating copies of a spreadsheet and culling rows based on
some condition (e.g., delete all rows whose cell in a certain column has
a value less than 100). Can optionally add dynamic Microsoft Excel
formulas afterward.
"""

from . import _constants

__version__ = _constants.__version__
__author__ = _constants.__author__
__contact__ = _constants.__email__
__website__ = _constants.__website__

from .xlsx_copycull import (
    add_formulas_to_column,
    WorkbookWrapper,
    WorksheetWrapper
)
