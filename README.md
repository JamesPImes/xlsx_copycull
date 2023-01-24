
# xlsx_copycull

A tool for streamlined copying of Microsoft Excel spreadsheets and
deleting only those rows that meet one or more user-specified
conditions, based on the value of a given cell or cells in each row.


## Background

Clients and coworkers often use Excel spreadsheets because they're more
intuitive than proper databases. So large spreadsheets commonly need to
be broken down into a series of smaller spreadsheets, with each copy
retaining only a subset of the original rows, and without destroying the
original spreadsheet. I wrote this module (building on
[openpyxl](https://pypi.org/project/openpyxl/)) to iteratively copy the
original 'master' spreadsheet and delete the unnecessary rows in each
copy, and then to add back in any necessary Excel formulas in each
remaining row.


## Table of Contents

* [To Install](#install)
* [Quick Example](#quick_example)
* [A more in-depth example](examples/generate_report_forms)
* ["How To"](#howto)
  * [Create a copy by initializing a `WorkbookWrapper` object](#copy)
  * [Cull unwanted rows with a `WorksheetWrapper` object](#cull)
  * [Subscript by sheet name](#subscripting)
  * [Access openpyxl objects with `.wb` and `.ws`](#openpyxl_objects)
  * [Add formulas to sheets](#formulas)
  * [Protect certain rows from modification or deletion](protected_rows)
  * [Save and close](#save_close)
  * [Use boolean operators `'OR'`, `'AND'`, `'XOR'`](#bool_oper)
  * [Reopen a closed ``WorkbookWrapper`` object](#reopen)
  * [Modify the original file without creating a copy](#no_copy)
* [Warnings](#warnings)


## <a name='install'>To install</a>

```
pip install git+https://github.com/JamesPImes/xlsx_copycull@master
```


## <a name='quick_example'>Quick Example</a>

Create a copy of `master1.xlsx`; and reduce the sheet named `'Sheet1'`
in that copy to only those rows that have a `'Price'` value of at least
`1000`.

```
from pathlib import Path
import xlsx_copycull

master_spreadsheet = Path(r'C:\Example\master1.xlsx')
spreadsheet_copy_fp = master_spreadsheet.parent / 'test_copy.xlsx'

# Copy the master spreadsheet to 'test_copy.xlsx' in the same directory.
wb_wrapper = xlsx_copycull.WorkbookWrapper(
    orig_fp=master_spreadsheet,
    copy_fp=spreadsheet_copy_fp)

# Open the copied .xlsx file.
wb_wrapper.load_wb()

# We'll modify existing 'Sheet1'.
ws_wrapper1 = wb_wrapper.stage_ws(
    ws_name='Sheet1',
    header_row=2)

# We'll keep only rows with a 'Price' value of at least 1000.
select_conditions={'Price': lambda price_val: price_val >= 1000}

ws_wrapper1.cull(select_conditions=select_conditions)
wb_wrapper.save_wb()
wb_wrapper.close_wb()
```

## A more in-depth example

Provided in `examples/generate_report_forms/` is a sample script I wrote
to showcase one way that I use this module in my consulting work:
splitting up a master review spreadsheet for teammates without
programming knowledge to research, fill out, and return.  You can
[read about it here](examples/generate_report_forms/readme.md) and
[see the script here](examples/generate_report_forms/script/generate_report_forms.py).


## <a name='howto'>"How To"</a>

Below is a series of guides for the main functionality of this module.


### <a name='copy'>Create a copy by initializing a `WorkbookWrapper` object</a>
Create a `WorkbookWrapper` object from the master spreadsheet, which
will automatically create a copy at the specified directory and filename
-- and the original spreadsheet will never again be touched by this
object or any subordinate objects.

```
from pathlib import Path
import xlsx_copycull

master_spreadsheet = Path(r'C:\Example\master1.xlsx')
spreadsheet_copy_fp = master_spreadsheet.parent / 'test_copy.xlsx'

# Copy the master spreadsheet to 'test_copy.xlsx'.
wb_wrapper = xlsx_copycull.WorkbookWrapper(
    orig_fp=master_spreadsheet,
    copy_fp=spreadsheet_copy_fp)

# Open the copy.
wb_wrapper.load_wb()
```

### <a name='cull'>Cull unwanted rows with a `WorksheetWrapper` object</a>

After creating the `WorkbookWrapper` object, stage a `WorksheetWrapper`
object with `.stage_ws()`. In this example, we'll set up the pre-existing
`'Sheet1'` for modification. (Note that this particular spreadsheet has
its headers in the second row.)

```
ws_wrapper1 = wb_wrapper.stage_ws(
    ws_name='Sheet1',
    header_row=2)
```

Then, prepare a dict whose keys specify which column headers to look
under, and each of whose values is a function to apply to the
corresponding cell value in each row, to determine whether to delete
that row.

```
select_conditions = {
    'Price': lambda cell_val: cell_val > 1000,
    'Color': lambda cell_val: cell_val in ('blue', 'red')
}
```

*The function need not be a lambda -- it can be any function that takes
a single argument (a given cell's value) and returns a bool or bool-like
value.*


Finally, we can pass this dict to `.cull()` to determine what rows get
deleted in `'Sheet1'`:

```
# Keep only those rows in 'Sheet1' that meet our critera. Specifically,
# wherever ANY specified condition is True, because we're using
# `bool_oper='OR'`
ws_wrapper1.cull(
    select_conditions=select_condititions,
    bool_oper='OR'
)

# Save and close.
wb_wrapper.save_wb()
wb_wrapper.close_wb()
```

### <a name='subscripting'>Subscripting by sheet name</a>

Once staged, we can access the `WorksheetWrapper` objects by
subscripting on the `WorkbookWrapper` object by sheet name if we need to:

```
wb_wrapper.stage_ws(ws_name='Sheet1', header_row=2)
ws_wrapper1 = wb_wrapper['Sheet1']
```

### <a name='openpyxl_objects'>Access openpyxl objects with `.wb` and `.ws`</a>

Once a `WorkbookWrapper` object has been created, access the underlying
openpyxl `Workbook` object with its `.wb` attribute:

```
wb_wrapper.wb  # an openpyxl Workbook object.
```

Similarly, once a `WorksheetWrapper` object is staged, we can access the
underlying openpyxl `Worksheet` object with its `.ws` attribute.

```
ws_wrapper1.ws  # an openpyxl Worksheet object.
```

*__Warning:__* When a `WorkbookWrapper` object is closed, its `.wb`
attribute is set to `None`; and the `.ws` attribute in each subordinate
`WorksheetWrapper` object is also set to `None`.  If it's reopened with
`wb_wrapper.load_wb()`, it will initialize *__new__* `Workbook` and
`Worksheet` objects at `.wb` and `.ws` respectively.

### <a name='formulas'>Add formulas to sheets</a>

Prepare a dict to generate Excel formulas. Specifically, the dict should
be keyed by column letter (`'A'`, `'Z'`, `'AA'`, whatever), and each
value should be a function (lambda or otherwise) that generates a
formula for each row.

In the following example, we'll write the formula `=(C3+D3)/$A$1` to
cell `B3`; `=(C4+D4)/$A$1` to cell `B4`, etc. And we'll write the
formula `=F1+F2` to cell `E3`; `=F2+F3` to cell `E4`; etc.

```
formulas_to_add = {
    "B": lambda row_num: f"=(C{row_num}+D{row_num})/$A$1",
    "E": lambda row_num: f"=F{row_num - 2}+F{row_num - 1}"
}

ws_wrapper1.add_formulas(formulas=formulas_to_add)
```

By default, the `.add_formulas()` method will apply to all unprotected
rows (*[see here](#protected_rows) for how to protect certain rows from
deletion or modification*).  But we can also choose to write formulas to
only certain rows with `rows=<list of ints>`.

```
# For 'Sheet2', write the same formulas as above, but only in
# rows 2 through 100.
ws_wrapper2.add_formulas(
    formulas=formulas_to_add,
    rows=range(2,101)
)
```
(Note that `rows=` can take any list-like object that contains ints.)


### <a name="protected_rows">Protect certain rows from modification or deletion</a>

When initially staging a worksheet with `.stage_ws()`, we can protect
certain rows from deletion with `protected_rows=<iterable of ints>`.

```
do_not_modify = [6, 7, 9] + list(range(14, 20))

ws_wrapper4 = wb_wrapper.stage_ws(
    ws_name='Sheet4',
    header_row=3,
    protected_rows=do_not_modify
)

# Rows 6, 7, 9, and 14 through 19 will not be deleted.
ws_wrapper4.cull(...)

# Those same rows will not have formulas added.
ws_wrapper4.add_formulas(...)
```

Note that the deletion of rows will trigger the automatic re-indexing of
protected rows to ensure the same rows are maintained. Access the
current indexing of rows that may not be modified in the
`.protected_rows` attribute.

Note also: If `protected_rows=` is NOT specified when initializing the
`WorksheetWrapper` object (most likely via `wb_wrapper.stage_ws(...)`),
then the protected rows will be all rows from the first through the
header row (i.e. the first __unprotected__ row will be the row
immediately following the headers).

Warning: `.protected_rows` will be protected from modification or
deletion ONLY by methods within this module. Those rows may be modified
directly, or by openpyxl itself, by other Python modules, etc.

Moreover, only the `.cull()` method will automatically reindex
`.protected_rows`. If inserting or deleting rows directly with the
openpyxl `Worksheet` object, you may need to manually adjust
`.protected_rows` accordingly.


### <a name='save_close'>Save and close</a>

Save and close the `WorkbookWrapper` object when we're done.

```
wb_wrapper.save_wb()
wb_wrapper.close_wb()

wb_wrapper.wb  # has now been set to `None`
ws_wrapper1.ws  # also `None`
ws_wrapper2.ws  # also `None`
```


We can [reopen the copied spreadsheet](#reopen) with
`wb_wrapper.load_wb()`, which will repopulate those `.wb` and `.ws`
attributes with *__new__* openpyxl `Workbook` and `Worksheet` objects.


### <a name='bool_oper'>Boolean operators `'OR'`, `'AND'`, `'XOR'`</a>

In the `.cull()` method, `bool_oper='AND'` is the default (i.e. keep
rows where *__all__* of the specified conditions are true).

If we wanted to keep rows where ANY condition is true (e.g., where
`'Price'` was greater than 1000, *__or__* where `'Color'` was
either `'blue'` or `'red'`), then we could pass `bool_oper='OR'`.

`bool_oper='XOR'` is also supported.

```
select_conditions = {
    'Price': lambda cell_val: cell_val > 1000,
    'Color': lambda cell_val: cell_val in ('blue', 'red')
    }

# Default functionality.
ws_wrapper1.cull(
    select_conditions=select_conditions,
    bool_oper='AND')

ws_wrapper2.cull(
    select_conditions=select_conditions,
    bool_oper='OR')

ws_wrapper3.cull(
    select_conditions=select_conditions,
    bool_oper='XOR')
```

It is currently not supported to mix and match bool operators in a
single pass (e.g., to delete row where conditions A and B are True, or
condition C is True).

### <a name='reopen'>Reopen a closed ``WorkbookWrapper`` object</a>

We can reopen the copied spreadsheet with `wb_wrapper.load_wb()`, which
will repopulate those `.wb` and `.ws` attributes with *__new__* openpyxl
`Workbook` and `Worksheet` objects.


### <a name='no_copy'>Modify the original file without creating a copy</a>

To modify a workbook without creating a copy, use the `no_copy=True`
parameter when initializing a `WorkbookWrapper`.

```
from pathlib import Path
import xlsx_copycull

master_spreadsheet = Path(r'C:\Example\master1.xlsx')

# Do NOT specify `copy_fp`.
wb_wrapper = xlsx_copycull.WorkbookWrapper(
    orig_fp=master_spreadsheet,
    no_copy=True)

# Open the ORIGINAL .xlsx file.
wb_wrapper.load_wb()

# ... etc.

# Save to the original file.
wb_wrapper.save_wb()

# Or save to a different file.
wb_wrapper.save_wb(master_spreadsheet.parent / 'test_copy.xlsx')
```


## <a name='warnings'>Warnings</a>

As with any script that uses openpyxl to modify spreadsheets, any
formulas that exist in the original spreadsheet will most likely NOT
survive the insertion or deletion of rows or columns (or changing of
worksheet names, etc.). Thus, it is highly recommended that you flatten
all existing formulas, and use the `.add_formulas()` method in the
`WorksheetWrapper` class -- or use the `copycull(<...>, formulas=<...>)`
function -- to the extent possible for your use case.

Also, you should familiarize yourself with the security warnings that
openpyxl gives [on their PyPI page](https://pypi.org/project/openpyxl/).
