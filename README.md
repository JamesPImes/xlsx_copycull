
# xlsx_copycull

A tool for streamlined copying of Microsoft Excel spreadsheets and deleting only those rows that meet a user-specified condition, based on the value of a given cell in each row.

Copyright © 2022, James P. Imes


## Background

I would often have large spreadsheets that needed to be broken up into a series of smaller spreadsheets, each with only a subset of the original rows. I wrote this module (building on [openpyxl](https://pypi.org/project/openpyxl/), obviously) to iteratively copy the original 'master' spreadsheet and delete the unnecessary rows in each copy, and then to add back in any necessary Excel formulas in each remaining row.


## Example 1 (extremely basic)

Copy a spreadsheet once, and retain only those rows whose `'Price'` value is an even number. Save the copy to the same directory as the original spreadsheet. (The first row of this spreadsheet contains headers, one of which is `'Price'`.)

```
from pathlib import Path
import xlsx_copycull

master_spreadsheet = Path(r'C:\Example\master1.xlsx')
copy_to_dir = master_spreadsheet.parent

delete_condition = lambda cell_val: cell_val % 2 == 1

xlsx_copycull.copy_cull_spreadsheet(
    wb_fp=master_spreadsheet,
    ws_name='Sheet1',
    match_col_name='Price',
    header_row=1,
    delete_condition=delete_condition,
    output_filename='new_copy.xlsx',
    copy_to_dir=copy_to_dir)
)
```

## Example 2

We have a spreadsheet that shows data on various products, including `'Price Per Unit'` (ranging from $1 to $300 per unit). We want to split this spreadsheet up into separate spreadsheets for $1 to $100, $101 to $200, and $201 to $300.

We also want to make sure that Column C contains a formula that adds each row's values in Columns D and F (i.e. the formula in cell `C2` is `=D2+F2`, cell `C3` is `=D3+F3`, etc.)

```
from pathlib import Path
import xlsx_copycull

master_spreadsheet = Path(r'C:\Example\master2.xlsx')

# This dict is keyed by column letter, and its value is a constructor
# for the appropriate Excel formula, based on the row number -- i.e.
# we'll add to each cell in Column C a formula to add cells D and F
# for that row (i.e. cell C2 will be '=D2+F2', etc.).
formulas_to_add = {
    "C": lambda row_num: "=D{0}+F{0}".format(row_num)
}

# Field we'll check against (i.e. the header of the column we'll look in).
field_to_check = 'Price Per Unit'

# We'll split up our spreadsheet into these ranges -- i.e. one spreadsheet 
# for $1/unit to $100/unit; another for $101/unit to $200/unit; etc.
split_mins_maxes = [(1, 100), (101, 200), (201, 300)]

for min_, max_ in split_mins_maxes:
    
    # The new filename of each copy.
    new_fn = f"example ({min_} - {max_}).xlsx"
    
    # Directory where we'll save each copy.
    copy_to_dir = master_spreadsheet.parent / 'splits'
    
    # Delete all rows where 'Price Per Unit' is outside the min and max.
    delete_condition = lambda cell_val: (cell_val < min_ or cell_val > max_)
    
    # Copy the spreadsheet, delete those rows in 'Sheet1' whose value
    # in 'Price Per Unit' meets the delete_condition, and then write
    # the formulas in the remaining cells in Column C. 
    xlsx_copycull.copy_cull_spreadsheet(
        wb_fp=master_spreadsheet,
        ws_name='Sheet1',
        match_col_name=field_to_check,
        header_row=1,
        delete_condition=delete_condition,
        output_filename=fn,
        copy_to_dir=copy_to_dir,
        formulas=formulas_to_add)
    )
```


## TODO

### # TODO: Basic guide to `WorkbookWrapper` class
### # TODO: Basic guide to `WorksheetWrapper` class


## Requirements

Python 3.6+
