# Copyright (c) 2021-2022, James P. Imes. All rights reserved.

"""
Tools for generating copies of a spreadsheet and culling rows based on
some condition (e.g., delete all rows whose cell in Column "A" has a
value less than 100). Can optionally add Microsoft Excel formulas
to the remaining rows afterward.
"""

from . import _constants

__version__ = _constants.__version__
__author__ = _constants.__author__
__contact__ = _constants.__email__
__website__ = _constants.__website__

from .xlsx_copycull import (
    copy_cull_spreadsheet,
    add_formulas_to_column,
    WorkbookWrapper,
    WorksheetWrapper
)