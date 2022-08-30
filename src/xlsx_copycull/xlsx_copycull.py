# Copyright (c) 2021-2022, James P. Imes. All rights reserved.

"""
Tools for generating copies of a spreadsheet and culling rows based on
some condition (e.g., delete all rows whose cell in a certain column has
a value less than 100). Can optionally add dynamic Microsoft Excel
formulas afterward.
"""

import os
import shutil
from pathlib import Path
import openpyxl


class WorkbookWrapper:
    """
    A wrapper class for openpyxl workbooks, with added methods for
    generating modified copies (e.g., reducing to only the rows relevant
    to a portion of the data).

    By design, this will leave the original spreadsheet alone and will
    generate a copy to modify.

    In particular, look into the ``.cull()`` and ``.add_formulas()``
    methods of the subordinate WorksheetWrapper objects (which objects
    are stored in the ``.ws_dict`` attribute).

    Before modifying any worksheet in the wrapped workbook with the
    added methods, you MUST stage it with the ``.stage_ws()`` method,
    specifying its name, the row containing the header (defaults to 1),
    and various optional parameters, such as which rows to leave alone
    (this will create a ``WorksheetWrapper`` object).

    Access the staged ``WorksheetWrapper`` objects either directly in
    the ``.ws_dict`` attribute (a dict, keyed by sheet name), or by
    subscripting on the ``WorkbookWrapper`` object (passing the sheet
    name):

        ``some_wb_wrapper.ws_dict['Sheet1'].cull(<...>)``

            ...is equivalent to...

        ``some_wb_wrapper['Sheet1'].cull(<...>)``

    (Remember, though, that worksheets must first be staged with
    ``.stage_ws()``, or this would raise a KeyError.)

    WARNING: As with any script that uses openpyxl to modify
    spreadsheets, any formulas that exist in the original spreadsheet
    will most likely NOT survive the insertion or deletion of rows or
    columns (or changing of worksheet names, etc.). Thus, it is highly
    recommended that you flatten all possible formulas, and use the
    `.add_formulas()` method in the WorksheetWrapper class to the extent
    possible for your use case.
    """

    def __init__(
            self,
            wb_fp: Path,
            output_filename: str,
            copy_to_dir: Path = None,
            uid=None):
        """
        A wrapper for an openpyxl Workbook object. Access the Workbook
        object directly in the ``.wb`` attribute.  The Workbook will
        be loaded at init.
        (NOTE: ``.wb`` will be set to ``None`` if the file is not
        currently open. Open it with the ``.load_wb()`` method, close it
        with ``.close_wb()`` (which will NOT save by default), and check
        whether it is currently open with the ``.is_loaded`` property.)

        :param wb_fp: Filepath to the workbook to load (and copy from).
        Must be in the .xlsx format!
        :param output_filename: The filename (NOT the full path) to
        which to save the copied workbook. Should include the '.xlsx'
        suffix.
        :param copy_to_dir: Directory in which to save the copied
        workbook. If not specified, will use the same directory as the
        base spreadsheet.
        :param uid: (Optional) An internal unique identifier.
        """
        # An internal identifier
        self.uid = uid

        # filepath to the original workbook
        self.wb_fp = Path(wb_fp)

        # a dict of subordinate WorksheetWrappers
        self.ws_dict = {}

        # The openpyxl workbook -- will be set to None whenever the wb
        # is NOT currently open.  (Check if this is currently set with
        # the property `self.is_loaded`)
        self.wb = None

        if copy_to_dir is None:
            copy_to_dir = self.wb_fp.parent

        self.copy_to_dir = Path(copy_to_dir)
        self.output_filename = Path(output_filename)
        self.new_fp = self.copy_to_dir / self.output_filename
        if self.wb_fp == self.new_fp:
            raise ValueError(
                "Error! The filepath created by combining `copy_to_dir` and "
                "`output_filename` may not be the same as the filepath to "
                "`wb_fp`!")

        self.copy_original()
        self.load_wb()

    @property
    def is_loaded(self):
        return self.wb is not None

    # ------------------------------------------------------------------
    # Subscriptable -- passes through keys to `.ws_dict` dict attribute.

    def __getitem__(self, item):
        try:
            return self.ws_dict[item]
        except KeyError:
            raise KeyError(
                f"worksheet '{item}' has not yet been staged (or "
                f"does not exist in this workbook). "
                f"Must first call `.stage_ws()`")

    # ------------------------------------------------------------------

    def stage_ws(
            self,
            ws_name,
            header_row: int = 1,
            first_modifiable_row: int = -1,
            protected_rows: set = None,
            rename_ws: str = None):
        """
        Prepare a worksheet for modification.

        :param ws_name: The (original) sheet name.
        :param header_row: The row containing headers (an int, indexed
        to 1)
        :param first_modifiable_row: (Optional) The first row that may
        be modified (an int, indexed to 1). If not set, will default to
        the first row after the `header_row`.
        :param protected_rows: A list-like object containing the rows
        that should never be deleted. Rows before `first_modifiable_row`
        and the header row will be automatically added.
        :param rename_ws: (Optional) A string, for how to rename the
        worksheet. Defaults to None, in which case, it will not be
        renamed. WARNING: If the worksheet is renamed, the new name will
        be the key for this worksheet, and NOT the original worksheet
        name.
        :return: The WorksheetWrapper object for the newly staged sheet
        (which is also stored to ``.ws_dict``, keyed by the sheet name).
        """
        wswp = WorksheetWrapper(
            wb_wrapper=self,
            ws_name=ws_name,
            header_row=header_row,
            protected_rows=protected_rows,
            first_modifiable_row=first_modifiable_row)
        self.ws_dict[ws_name] = wswp
        if rename_ws:
            self.rename_ws(ws_name, new_name=rename_ws)
        return wswp

    def copy_original(self, fp=None, stage_new_fp=False):
        """
        Copy the source spreadsheet to the new filepath at `fp`, and
        store that new filepath to `self.new_fp`. (If `fp` is not
        specified here, will default to whatever is already set in
        `self.new_fp`.)
        :param fp: The filepath to copy to.
        :param stage_new_fp: A bool, whether to set the filepath of the
        newly copied workbook as the target workbook of this
        WorkbookWrapper object. That is, whether or not the newly copied
        spreadsheet is the one we want to be working on. Defaults to
        False.
        (NOTE: If the workbook is currently open and
        ``stage_new_fp=True`` is passed, it will raise a RuntimeError.
        To avoid that error, save and close the workbook first with
        ``.close_wb()``.)
        :return: None
        """
        if self.is_loaded and stage_new_fp:
            raise RuntimeError(
                "Workbook is currently open. Save and close with "
                "`.close_wb()` before copying.")
        if fp is None:
            fp = self.new_fp
        fp = Path(fp)
        os.makedirs(fp.parent, exist_ok=True)
        shutil.copy(self.wb_fp, fp)
        self.new_fp = fp
        return None

    def delete_ws(self, ws_name):
        """
        Delete a worksheet from the workbook. (The worksheet need not be
        staged.)
        :param ws_name: The name of the worksheet to discard.
        :return: None
        """
        self.mandate_loaded()
        self.wb.remove(self.wb[ws_name])
        if ws_name in self.ws_dict.keys():
            self.ws_dict.pop(ws_name)
        return None

    def load_wb(self):
        """
        Open the workbook at the filepath stored in `self.new_fp` (and
        behind the scenes, inform all subordinate worksheets that they
        are now open for modification -- by setting their `.ws`
        attributes to the appropriate openpyxl worksheet object.)
        :return: None
        """
        if self.is_loaded:
            return
        self.wb = openpyxl.load_workbook(self.new_fp)
        # Update all of the staged worksheets.
        self._inform_subordinates()
        return None

    def _inform_subordinates(self):
        """
        INTERNAL USE:

        Inform the subordinate WorksheetWrapper objects whether the
        workbook has been opened or closed. If opened, set their `.ws`
        attributes to their respective openpyxl worksheet.
        :return:
        """
        is_loaded = self.is_loaded
        for ws_name, ws_wrapper in self.ws_dict.items():
            ws = None
            if is_loaded:
                ws = self.wb[ws_name]
            ws_wrapper.ws = ws
        return None

    def close_wb(self, save=False):
        """
        Close the workbook, and inform the subordinates that they cannot
        be modified until the workbook is reopened with ``.load_wb()``.
        :param save: Whether to save before closing. (False by default.)
        :return: None
        """
        if not self.is_loaded:
            return None
        if save:
            self.save_wb()
        self.wb.close()
        self.wb = None
        # Update all of the staged worksheets.
        self._inform_subordinates()
        return None

    def save_wb(self):
        self.mandate_loaded()
        self.wb.save(self.new_fp)
        return None

    def mandate_loaded(self):
        """Raise an error if the wb is not currently loaded."""
        if not self.is_loaded:
            raise RuntimeError("Workbook is not currently open")
        return None

    def rename_ws(self, old_name, new_name):
        """
        Rename a worksheet. (Workbook must be open, and worksheet with
        `old_name` must already be staged.)

        Note that renaming the worksheet will also modify the
        corresponding ``.ws_dict`` key:

            ```
            ws_wrapper1 = wb_wrapper.ws_dict['Sheet1']  # OK
            ws_wrapper1 = wb_wrapper['Sheet1']  # OK
            wb_wrapper_obj.rename_ws('Sheet1', 'Prices')
            ws_wrapper1 = wb_wrapper.ws_dict['Prices']  # new sheet name
            ws_wrapper1 = wb_wrapper['Prices']  # new sheet name
            ws_wrapper1 = wb_wrapper['Sheet1']  # raises KeyError.
            ```
        """
        self.mandate_loaded()
        self.ws_dict[old_name].rename_ws(new_name=new_name)
        return None


class WorksheetWrapper:
    """
    A wrapper class for openpyxl worksheets, with added methods for
    culling rows (based on whether cell values match a specified
    condition) and adding formulas.
    """
    def __init__(
            self,
            wb_wrapper: WorkbookWrapper,
            ws_name: str,
            header_row: int = 1,
            protected_rows: set = None,
            first_modifiable_row: int = -1):
        """
        :param wb_wrapper: The parent WorkbookWrapper object.
        :param ws_name: The name of this worksheet.
        :param header_row: The row containing headers (an int, indexed
        to 1)
        :param protected_rows: (Optional) A list-like object containing
        the rows that should never be modified or deleted. Rows before
        ``first_modifiable_row`` and the header row will be
        automatically added. (NOTE: `protected_rows` may change behind
        the scenes if rows are deleted by ``.cull()``.)
        :param first_modifiable_row: (Optional) The first row that may
        be modified or deleted (an int, indexed to 1). If not set, will
        default to the first row after the `header_row`.
        """
        self.wb_wrapper = wb_wrapper
        self.ws_name = ws_name
        self.ws = None
        if wb_wrapper.is_loaded:
            self.ws = wb_wrapper.wb[ws_name]
        self.header_row = header_row
        self.protected_rows = self._populate_protected_rows(
            protected_rows, first_modifiable_row)
        self.last_protected_rows = self.protected_rows

    @property
    def is_loaded(self):
        return self.ws is not None

    def mandate_loaded(self):
        """Raise an error if the wb is not currently loaded."""
        if not self.is_loaded:
            raise RuntimeError("Workbook is not currently open")
        return None

    def load(self):
        self.wb_wrapper.load_wb()
        return None

    def save(self):
        self.mandate_loaded()
        self.wb_wrapper.save_wb()
        return None

    def _populate_protected_rows(
            self, explicitly_protected, first_modifiable_row):
        """
        INTERNAL USE:

        Lock down which rows may never be deleted.
        """
        header_row = self.header_row
        if first_modifiable_row <= 0:
            first_modifiable_row = header_row + 1

        protected_rows = set()
        if explicitly_protected is not None:
            protected_rows = set(explicitly_protected)

        protected_rows.update(set(range(1, first_modifiable_row)))
        protected_rows.add(header_row)  # Never delete the header.
        return protected_rows

    def rename_ws(self, new_name):
        """
        Rename this worksheet.

        :param new_name: The new name for this sheet.
        :return: None
        """
        self.mandate_loaded()
        old_name = self.ws_name
        self.ws.title = new_name
        self.ws_name = new_name
        # Update the parent WBWrapper with the new sheet name, and
        # discard the old name.
        self.wb_wrapper.ws_dict[new_name] = self.wb_wrapper.ws_dict.pop(old_name)
        return None

    def cull(
            self,
            delete_conditions: dict,
            bool_oper='AND',
            protected_rows=None):
        """
        Cull the spreadsheet, based on the ``delete_conditions``.  If
        more than one delete_condition is used, specify whether to apply
        'AND', 'OR', or 'XOR' boolean logic to the resulting sets by
        passing one of those as ``bool_oper`` (defaults to 'AND').

        NOTE: ``protected_rows`` is a list (or set) of integers, being
        the row numbers for those rows that should NEVER be deleted
        (indexed to 1). If rows above those numbers get deleted, the
        resulting indexes in ``protected_rows`` would not be accurate,
        so this method adjusts the protected rows to their new position.
        The resulting row numbers are stored to attribute
        ``.last_protected_rows``.  IF AND ONLY IF ``protected_rows`` is
        NOT passed as an arg here, it will be pulled from
        ``.protected_rows``, and the resulting ``protected_rows`` will
        ALSO be stored to ``.protected_rows`` (in addition to
        ``.last_protected_rows``).

        (The reason for this design choice was that rows may need to be
        protected for one operation, but not for another -- but we want
        to be able to track any changes to them. This way, they can be
        accessed in ``.protected_rows`` and/or ``.last_protected_rows``
        after calling ``.cull()`` but before calling the next method,
        which might change them.)

        :param delete_conditions: A dict of column_header-to-
        delete_condition pairs, to determine which rows should be
        deleted. Specifically, keyed by the header of the column to
        check under, and whose value is a function to be applied to the
        value of the cell under that column, which returns a bool. If
        the function returns True (or a True-like value) when applied to
        the cell's value, that row will be marked for deletion.

        :param bool_oper: When using more than one delete conditions
        (i.e. more than one key in the dict), use this to determine
        whether to apply OR, AND, or XOR to the resulting rows to be
        deleted.  Pass one of the following:  'AND', 'OR', 'XOR'.
        (Defaults to 'AND'.)

        :param protected_rows: (Optional) A list-like object containing
        the rows that should never be deleted. If not specified here,
        will pull from what is set in ``.protected_rows``.  (See
        comments above regarding ``.protected_rows`` and
        ``.last_protected_rows``.)

        :return: None
        """
        self.mandate_loaded()

        if not delete_conditions:
            return None

        store_protected_rows = False
        if protected_rows is None:
            protected_rows = self.protected_rows
            store_protected_rows = True

        header_row = self.header_row

        ws = self.ws
        all_to_delete = []
        # Apply each delete condition to the appropriate column.
        for field, delete_condition in delete_conditions.items():
            match_col = self.find_match_col(header_row, field)

            # Mark for deletion all those rows that match our criteria.
            # Convert to a set and add it to the list.
            to_delete = (
                j for j in range(1, ws.max_row + 1)
                if (j not in self.protected_rows
                    and delete_condition(ws.cell(row=j, column=match_col).value)
                    )
            )
            all_to_delete.append(set(to_delete))

        # Apply the boolean operator to determine which rows to delete.
        final_to_delete = self._apply_bool_operator(all_to_delete, bool_oper)

        # convert our raw to_delete list down to a list of 2-tuples (ranges,
        # inclusive of min/max)
        rges = find_ranges(final_to_delete)

        # Delete ranges of rows from bottom-up.
        rges.reverse()

        for rge in rges:
            row = rge[0]
            num_rows_to_delete = rge[1] - rge[0] + 1
            ws.delete_rows(row, num_rows_to_delete)

        # Adjust any protected row numbers upward, if any higher rows
        # were deleted
        protected_rows_after_cull = []
        for rn in protected_rows:
            for rge in rges:
                if rge[1] < rn:
                    rn -= (rge[1] - rge[0] + 1)
            protected_rows_after_cull.append(rn)

        new_protected_rows = set(protected_rows_after_cull)
        self.last_protected_rows = new_protected_rows
        if store_protected_rows:
            self.protected_rows = new_protected_rows

        return None

    def find_match_col(self, header_row, match_col_name):
        """
        Find the match column number, based on its header name.
        NOTE: Will return the first match, so avoid duplicate header
        names.
        """

        self.mandate_loaded()
        ws = self.ws

        try:
            match_col = [
                c.column for c in ws[header_row] if c.value == match_col_name][0]
        except IndexError:
            raise RuntimeError(
                f"ERROR! Could not find the column name {match_col_name!r} "
                f"in header row ({header_row}).")
        return match_col

    def modifiable_rows(self, protected_rows=None) -> list:
        """
        Get a list of row numbers that currently exist in the
        spreadsheet (indexed to 1), and which are NOT in
        `protected_rows`.

        :param protected_rows: (Optional) A list-like object containing
        the rows that should never be deleted. If not specified here,
        will pull from what is set in ``.protected_rows``.
        """
        self.mandate_loaded()
        if protected_rows is None:
            protected_rows = self.protected_rows
        if protected_rows is None:
            protected_rows = set()
        row_nums = [
            j for j in range(1, self.ws.max_row + 1)
            if j not in protected_rows
        ]
        return row_nums

    def add_formulas(self, formulas, rows=None, protected_rows=None):
        """
        Add formulas to the working spreadsheet in `self.ws`.

        :param formulas: A dict keyed by column name (i.e. "E") whose
        values are a function that generates the formula, based on the
        row number.

        For example, to generate these formulas...

            Column R --> =F5*AB5/$S$1  (for an example row 5)
            Column S --> =AB5*AC5  (for an example row 5)

            ...we would pass this dict:

            formulas = {
                "R": lambda row_num: "=F{0}*AB{0}/$S$1".format(row_num),
                "S": lambda row_num: "=AB{0}*AC{0}".format(row_num)
            }

        :param rows: The rows where formulas should be added. If
        ``rows`` is specified here, it will IGNORE ``protected_rows``
        (potentially adding formulas to all rows in ``rows``, even if
        they are also in ``protected_rows``).  However, if ``rows`` is
        NOT specified, it will insert a formula into every row EXCEPT
        ``protected_rows``.

        :param protected_rows: (Optional) A list-like object containing
        the rows that should never be deleted. If not specified here,
        will pull from what is set in ``.protected_rows``. NOTE: If
        ``.cull()`` was called before this method, then the
        ``protected_rows`` may have changed since this object was
        initialized. See comments under ``.cull()`` for a more complete
        discussion.

        :return: None
        """
        self.mandate_loaded()
        if formulas is None:
            return None
        if rows is None:
            rows = self.modifiable_rows(protected_rows=protected_rows)
        for column, formula in formulas.items():
            add_formulas_to_column(
                ws=self.ws, column=column, rows=rows, formula=formula)
        self.last_protected_rows = protected_rows
        return None

    @staticmethod
    def _apply_bool_operator(list_of_sets: list, operator: str) -> set:
        """
        INTERNAL USE:
        Apply the specified boolean operator to a list of sets.

        :param list_of_sets: A list of sets.
        :param operator: Which boolean operator to apply -- either
        'AND', 'OR', or 'XOR'.
        :return: A set of the resulting elements.
        """
        operator = operator.upper()
        final = set()
        if list_of_sets:
            final = list_of_sets.pop()
        if operator == 'OR':
            for new_set in list_of_sets:
                final.update(new_set)
        elif operator == 'AND':
            for new_set in list_of_sets:
                final.intersection_update(new_set)
        elif operator == 'XOR':
            for new_set in list_of_sets:
                final = final ^ new_set
        else:
            raise ValueError(
                f"`operator` must be one of ['OR', 'AND', 'XOR']. "
                f"Passed {operator!r}")
        return final


def copycull(
        wb_fp: Path,
        ws_name: str,
        header_row: int,
        delete_conditions: dict,
        output_filename: str,
        bool_oper: str = 'AND',
        copy_to_dir: Path = None,
        first_modifiable_row: int = -1,
        protected_rows=None,
        rename_ws=None,
        formulas=None):
    """
    Copy a target spreadsheet and cull the copy down to only those rows
    that do NOT match the ``delete_conditions``.  Optionally add Excel
    formulas with ``formulas=<dict>`` (see below).

    (A function that combines the basic functionality of WorkbookWrapper
    and WorksheetWrapper objects. Returns the WorkbookWrapper object.)

    :param wb_fp: The filepath to the source .xlsx file.

    :param ws_name: The name of the worksheet within that workbook to
    cull.

    :param header_row: The row containing headers (an int, indexed to 1)

    :param first_modifiable_row: (Optional) The first row that may be
    deleted (an int, indexed to 1). If not set, will default to the
    first row after the `header_row`.

    :param delete_conditions: A dict of column_header-to-
    delete_condition pairs, to determine which rows should be deleted.
    Specifically, keyed by the header of the column to check under, and
    whose value is a function to be applied to the value of the cell
    under that column, which returns a bool. If the function returns
    True (or a True-like value) when applied to the cell's value, that
    row will be marked for deletion.

    :param bool_oper: When using more than one delete conditions
    (i.e. more than one key in the dict), use this to determine
    whether to apply OR, AND, or XOR to the resulting rows to be
    deleted.  Pass one of the following:  'AND', 'OR', 'XOR'.
    (Defaults to 'AND'.)

    :param output_filename: The filename (NOT the full path) to which to
    the copied workbook. Should include the '.xlsx' suffix.

    :param copy_to_dir: Directory in which to save the copied workbook.
    If not specified, will use the same directory as the base
    spreadsheet.

    :param protected_rows: (Optional) A list-like object containing the
    rows that should never be deleted. Rows before `first_deletable_row`
    and the header row will be automatically added.

    :param rename_ws: (Optional) A string, for how to rename the
    modified worksheet. Defaults to None, in which case, it will not be
    renamed. WARNING: Using this feature will prevent `sanity_check()`
    from working.

    :param formulas: (Optional) A dict keyed by column name (i.e. "E")
    whose values are a function that generates the formula, based on the
    row number -- such as:
        Column R --> =F5*AB5/$S$1  (for an example row 5)
        Column S --> =AB5*AC5  (for an example row 5)
    formulas = {
        "R": lambda row_num: "=F{0}*AB{0}/$S$1".format(row_num),
        "S": lambda row_num: "=AB{0}*AC{0}".format(row_num)
    }
    :return: A WorkbookWrapper object. (Access its ``.new_fp`` for the
    filepath to the copied spreadsheet.)
    """

    # Wrap, copy, and load our workbook
    wbwp = WorkbookWrapper(
        wb_fp=wb_fp, output_filename=output_filename, copy_to_dir=copy_to_dir)

    # Stage our worksheet, and grab the resulting WorksheetWrapper obj.
    wswp = wbwp.stage_ws(
        ws_name=ws_name,
        header_row=header_row,
        first_modifiable_row=first_modifiable_row,
        protected_rows=protected_rows,
        rename_ws=rename_ws)

    # Cull down to the desired rows.
    wswp.cull(delete_conditions=delete_conditions, bool_oper=bool_oper)

    # Add the requested formulas (if any).
    wswp.add_formulas(formulas=formulas)

    # Close and save our workbook.
    wbwp.close_wb(save=True)

    # Return the WorkbookWrapper.
    return wbwp


def find_ranges(nums: set) -> list:
    """
    Find ranges of consecutive integers in the set. Returns a list of
    tuples, of the first and last numbers (inclusive) in each sequence.
    :param nums: A set of integers.
    :return: A list of 2-tuples of integers, being the min and max
    of each range (inclusive).
    """
    nnums = list(nums)
    nnums.sort()
    starts = [n for n in nnums if n - 1 not in nums]
    ends = [n for n in nnums if n + 1 not in nums]
    return [*zip(starts, ends)]


def add_formulas_to_column(ws, column: str, rows: list, formula):
    """
    Add formulas to every row in a column in the worksheet, based on
    each row number.

    :param ws: An openpyxl worksheet.
    :param column: The column's letter name (i.e. "D").
    :param rows: A list of integers, being the row numbers where to add
    the formula. (Indexed to 1)
    :param formula: A function that will generate a formula based on the
    current row number -- such as:
        lambda row_num: "=F{0}*AB{0}/$S$1".format(row_num)
            --> Generates '=F5*AB5/$S$1'   (for an example row 5).
    :return: None.
    """
    for row in rows:
        cell_name = f"{column}{row}"
        ws[cell_name] = formula(row)
        ws[cell_name].number_format = "General"
    return None


__all__ = [
    copycull,
    add_formulas_to_column,
    WorkbookWrapper,
    WorksheetWrapper
]
