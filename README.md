
# xlsx_copycull

A tool for streamlined copying of Microsoft Excel spreadsheets and deleting only those rows that meet a user-specified condition, based on the value of a given cell in each row.

Copyright © 2022, James P. Imes


## Background

Clients and coworkers often use Excel spreadsheets because they're more intuitive than proper databases. So I often have large spreadsheets that need to be broken down into a series of smaller spreadsheets, each copy retaining only a subset of the original rows (without destroying the original spreadsheet). I wrote this module (building on [openpyxl](https://pypi.org/project/openpyxl/), obviously) to iteratively copy the original 'master' spreadsheet and delete the unnecessary rows in each copy, and then to add back in any necessary Excel formulas in each remaining row.


## How it works

Basically, we prepare a dict whose keys specify which column headers to look under, and each of whose values is a function to apply to the corresponding cell value in each row, to determine whether to delete a given row. Like so:

```
delete_conditions = {
    'Price': lambda cell_val: cell_val <= 1000,
    'Color': lambda cell_val: cell_val not in ('blue', 'red')
    }
```

We can pass this to the simplified `xlsx_copycull.copy_cull_spreadsheet()` function (with other required arguments) -- see Examples 1 and 2 below.

Or we can use it with `WorkbookWrapper` and `WorksheetWrapper` objects if we need to do other tasks with the underlying `Workbook` or `Worksheet` objects.  (# TODO: Additional write-up on these two wrapper classes.)


## Example 1 (basic)

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

xlsx_copycull.copy_cull_spreadsheet(
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


## Example 2

We have a spreadsheet that shows data on various products, including `'Price Per Unit'` (ranging from $1 to $300 per unit). We want to split this spreadsheet up into separate spreadsheets for $1 to $100, $101 to $200, and $201 to $300.

We also want to make sure that under Column C, each row contains a formula that adds that same row's values in Columns D and F and divides by the value in cell `A1` (i.e. the formula in cell `C2` should be `=(D2+F2)/$A$1`, cell `C3` should be `=(D3+F3)/$A$1`, etc.)

```
from pathlib import Path
import xlsx_copycull

master_spreadsheet = Path(r'C:\Example\master2.xlsx')

# This dict is keyed by column letter, and its value is a constructor
# for the appropriate Excel formula, based on the row number -- i.e.
# we'll add to each cell in Column C a formula to add cells D and F
# for that row (i.e. cell C2 will be '=D2+F2', etc.).
formulas_to_add = {
    "C": lambda row_num: "=(D{0}+F{0})/$A$1".format(row_num)
}

# We'll split up our spreadsheet into these ranges -- i.e. one spreadsheet 
# for $1/unit to $100/unit; another for $101/unit to $200/unit; etc.
split_mins_maxes = [(1, 100), (101, 200), (201, 300)]

for min_, max_ in split_mins_maxes:
    
    # The new filename of each copy.
    new_fn = f"example (ppu {min_} - {max_}).xlsx"
    
    # Directory where we'll save each copy.
    copy_to_dir = master_spreadsheet.parent / 'splits'
    
    # Delete all rows where 'Price Per Unit' is outside the min and max.
    # (Checks the value under the 'Price Per Unit' column against the 
    # lambda function.)
    delete_conditions = {
        'Price Per Unit': lambda cell_val: (cell_val < min_ or cell_val > max_)
        }
    
    # Copy the spreadsheet, delete those rows in 'Sheet1' whose value
    # in 'Price Per Unit' meets the delete_condition, and then write
    # the formulas in the remaining cells in Column C. 
    xlsx_copycull.copy_cull_spreadsheet(
        wb_fp=master_spreadsheet,
        ws_name='Sheet1',
        header_row=1,
        delete_conditions=delete_conditions,
        output_filename=fn,
        copy_to_dir=copy_to_dir,
        formulas=formulas_to_add)
    )
```


## Warnings

As with any script that uses openpyxl to modify spreadsheets, any formulas that exist in the original spreadsheet will most likely NOT survive the insertion or deletion of rows or columns (or changing of worksheet names, etc.). Thus, it is highly recommended that you flatten all existing formulas, and use the `.add_formulas()` method in the `WorksheetWrapper` class -- or the `formulas=<>` kwarg in the `copy_cull_spreadsheet()` function -- to the extent possible for your use case.

Also, you should familiarize yourself with the security warnings that openpyxl gives [on their PyPI page](https://pypi.org/project/openpyxl/).


## TODO

### # TODO: Basic guide to `WorkbookWrapper` class
### # TODO: Basic guide to `WorksheetWrapper` class


## Requirements

Python 3.6+
