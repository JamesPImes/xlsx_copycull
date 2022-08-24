
# xlsx_copycull

A tool for streamlined copying of Microsoft Excel spreadsheets and deleting only those rows that meet one or more user-specified conditions, based on the value of a given cell or cells in each row.


## Background

Clients and coworkers often use Excel spreadsheets because they're more intuitive than proper databases. So large spreadsheets commonly need to be broken down into a series of smaller spreadsheets, with each copy retaining only a subset of the original rows, and without destroying the original spreadsheet. I wrote this module (building on [openpyxl](https://pypi.org/project/openpyxl/), obviously) to iteratively copy the original 'master' spreadsheet and delete the unnecessary rows in each copy, and then to add back in any necessary Excel formulas in each remaining row.


## How to use

Basically, we prepare a dict whose keys specify which column headers to look under, and each of whose values is a function to apply to the corresponding cell value in each row, to determine whether to delete that row. Like so:

```
delete_conditions = {
    'Price': lambda cell_val: cell_val <= 1000,
    'Color': lambda cell_val: cell_val not in ('blue', 'red')
    }
```

*The function need not be a lambda -- it can be any function that takes a single argument (a given cell's value) and returns a bool or bool-like value.*

We can pass this dict to the simplified `xlsx_copycull.copycull()` function (with other required arguments) -- see [Example 1](#example1) and [Example 2](#example2) below.

Or we can [use it with `WorkbookWrapper` and `WorksheetWrapper` objects](#wbws) if we need to cull rows in multiple sheets within the workbook, or do other tasks with the underlying `Workbook` or `Worksheet` objects.


### <a name="example1">Example 1</a> - `copycull()`

Copy a spreadsheet only a single time, and retain only those rows whose `'Price'` value is greater than 1000 *__and__* whose `'Color'` value is `'blue'` or `'red'`. Save the copy to the same directory as the original spreadsheet.

Note: This spreadsheet contains headers in the first row (`header_row=1`), among which are `'Price'` and `'Color'`.

```
from pathlib import Path
import xlsx_copycull

master_spreadsheet = Path(r'C:\Example\master1.xlsx')
copy_to_dir = master_spreadsheet.parent

delete_conditions = {
    'Price': lambda cell_val: cell_val <= 1000,
    'Color': lambda cell_val: cell_val not in ('blue', 'red')
    }

xlsx_copycull.copycull(
    wb_fp=master_spreadsheet,
    ws_name='Sheet1',
    header_row=1,
    delete_conditions=delete_conditions,
    bool_oper='AND',
    output_filename='new_copy.xlsx',
    copy_to_dir=copy_to_dir)
)
```

NOTE: `bool_oper='AND'` is the default (i.e. delete rows where all of the specified conditions are true).

If we wanted to delete rows where the `'Price'` was less than or equal to 1000, *__or__* where `'Color'` was neither `'blue'` nor `'red'`, then we could pass `bool_oper='OR'`.

`bool_oper='XOR'` is also supported.

It is currently not possible to mix and match bool operators in a single pass (e.g., to delete row where conditions A and B are True, or condition C is True).


### <a name="example2">Example 2</a> - generate multiple copies with `copycull()` and add Excel formulas

We have a spreadsheet that shows data on various products, including `'Price Per Unit'` (ranging from $1 to $300 per unit). We want to split this spreadsheet up into separate spreadsheets for $1 to $100, $100 to $200, and $200 to $300.

We also want to make sure that under Column C, each row contains a formula that adds that same row's values in Columns D and F and divides by the value in cell `A1` (i.e. the formula in cell `C2` should be `=(D2+F2)/$A$1`, cell `C3` should be `=(D3+F3)/$A$1`, etc.)

```
from pathlib import Path
import xlsx_copycull

master_spreadsheet = Path(r'C:\Example\master2.xlsx')

# This dict is keyed by column letter, and the value is a constructor
# for the appropriate Excel formula.
formulas_to_add = {
    "C": lambda row_num: "=(D{0}+F{0})/$A$1".format(row_num)
}

# We'll split up our spreadsheet into these ranges -- i.e. one spreadsheet 
# for $1/unit to $100/unit; another for $100/unit to $200/unit; etc.
split_mins_maxes = [(1, 100), (100, 200), (200, 300)]

for min_, max_ in split_mins_maxes:
    
    # The new filename of each copy.
    new_fn = f"example (ppu {min_} - {max_}).xlsx"
    
    # Directory where we'll save each copy.
    copy_to_dir = master_spreadsheet.parent / 'splits'
    
    # Delete all rows where 'Price Per Unit' is outside the min and max.
    # (Checks the value under the 'Price Per Unit' column against the 
    # lambda function.)
    delete_conditions = {
        'Price Per Unit': lambda cell_val: (cell_val <= min_ or cell_val > max_)
        }
    
    # Copy the spreadsheet, delete those rows in 'Sheet1' whose value
    # in 'Price Per Unit' meets the delete_condition, and then write
    # the formulas in the remaining cells in Column C. 
    xlsx_copycull.copycull(
        wb_fp=master_spreadsheet,
        ws_name='Sheet1',
        header_row=1,
        delete_conditions=delete_conditions,
        output_filename=fn,
        copy_to_dir=copy_to_dir,
        formulas=formulas_to_add)
    )
```

### <a name="wbws">Have more control with `WorkbookWrapper` and `WorksheetWrapper` objects</a>

Creating a `WorkbookWrapper` object automatically copies the master spreadsheet and opens the copy. The original spreadsheet is never touched after the `WorkbookWrapper` is created.

```
from pathlib import Path
from xlsx_copycull import WorkbookWrapper, WorksheetWrapper

master_spreadsheet = Path(r"C:\Example\master3.xlsx")

wb_wrapper = WorkbookWrapper(
    wb_fp=master_spreadsheet,
    output_filename='test_copy.xlsx',
    copy_to_dir=master_spreadsheet.parent
)

wb_wrapper.wb  # openpyxl Workbook object of the copied workbook.
```

Populate a `WorksheetWrapper` for each existing sheet we want to modify.

```
# Continuing the above codeblock...

ws_wrapper1 = wb_wrapper.stage_ws(
    ws_name='Sheet1',
    header_row=2
)
ws_wrapper2 = wb_wrapper.stage_ws(
    ws_name='Sheet2',
    header_row=3
)

ws_wrapper1.ws  # openpyxl Worksheet object of 'Sheet1' in the copied workbook
ws_wrapper2.ws  # openpyxl Worksheet object of 'Sheet2' in the copied workbook
```

Once staged, we can access the `WorksheetWrapper` objects by subscripting by sheet name if we need to:

```
ws_wrapper1 = wb_wrapper['Sheet1']
ws_wrapper2 = wb_wrapper['Sheet2']
```

Now we can cull each sheet by specifying the appropriate `delete_conditions`.  Each time `.cull()` is called, the copied workbook will save unless you specify `.cull(<...>, save=False)`.

```
del_condits = {
    'Price': lambda cell_val: cell_val < 1000,
    'Color': lambda cell_val: cell_val not in ('blue', 'red')
}

ws_wrapper1.cull(
    delete_conditions=del_condits,
    bool_oper='OR'
)

ws_wrapper2.cull(
    delete_conditions=del_condits,
    bool_oper='AND'
}
```

We can also add formulas to the sheets. (This will also save the copied workbook unless we specify `save=False`.)

By default, will apply to all unprotected rows**.  But we can also choose to write formulas to only certain rows with `rows=<list of ints>`.

** *(All rows below the header are unprotected by default, unless specified otherwise by the user when staging or initializing the worksheet -- reference the docstrings for `WorkbookWrapper.stage_ws` and `WorksheetWrapper.__init__` if you need this functionality).*

```
formulas_to_add = {
    "B": lambda row_num: "=(C{0}+D{0})/$A$1".format(row_num),
    "E": lambda row_num: "=F{0}+F{1}".format(row_num - 2, row_num - 1)
}

ws_wrapper1.add_formulas(formulas=formulas_to_add)

# For 'Sheet2', write formulas only in rows number 2 to 100.
rows_for_formulas = list(range(2,101))

ws_wrapper2.add_formulas(
    formulas=formulas_to_add,
    rows=rows_for_formulas
)
```

Close the `WorkbookWrapper` object when we're done. (Closing will also save unless `save=False` is passed.)

```
wb_wrapper.close_wb()

wb_wrapper.wb  # has now been set to `None`
ws_wrapper1.ws  # also `None`
ws_wrapper2.ws  # also `None`
```

We can reopen the copied spreadsheet with `wb_wrapper.open()`, which will repopulate those `.wb` and `.ws` attributes with new openpyxl `Workbook` and `Worksheet` objects.


## Warnings

As with any script that uses openpyxl to modify spreadsheets, any formulas that exist in the original spreadsheet will most likely NOT survive the insertion or deletion of rows or columns (or changing of worksheet names, etc.). Thus, it is highly recommended that you flatten all existing formulas, and use the `.add_formulas()` method in the `WorksheetWrapper` class -- or use the `copycull(<...>, formulas=<...>)` function -- to the extent possible for your use case.

Also, you should familiarize yourself with the security warnings that openpyxl gives [on their PyPI page](https://pypi.org/project/openpyxl/).


## Requirements

Python 3.6+
