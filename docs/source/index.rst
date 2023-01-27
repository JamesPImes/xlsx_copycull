
xlsx_copycull
=============

.. automodule:: xlsx_copycull


Quick example
-------------

.. code-block:: python

    from pathlib import Path
    import xlsx_copycull

    # The original base spreadsheet will not get modified.
    base_spreadsheet_fp = Path(r"some\path\original_spreadsheet.xlsx")
    copy_to_fp = Path(r"some\other\path\copied_spreadsheet.xlsx")

    wb_wrapper = xlsx_copycull.WorkbookWrapper(
        orig_fp=base_spreadsheet_fp,
        copy_fp=copy_to_fp)
    wb_wrapper.load_wb()  # loads the copied .xlsx file

    # In 'Sheet1', keep only rows with a 'Price' value of at least 10,
    # and whose 'Team Name' value is either "ABC" or "XYZ".
    sheet_wrapper = wb_wrapper.stage_ws('Sheet1', header_row=1)
    select_conditions = {
        "Price": lambda price_val: price_val >= 10,
        "Team Name": lambda team_name: team_name in ["ABC", "XYZ"]
    }
    sheet_wrapper.cull(select_conditions=select_conditions)

    # Add a formula to each cell in Column G below the header.
    # Here, Cell G2 will be "=A2+B2"; G3 will be "=A3+B3"; etc.
    formulas_to_add = {
        "G": lambda row_num: f"=A{row_num}+B{row_num}"
    }
    sheet_wrapper.add_formulas(formulas=formulas_to_add)

    # Access the wrapped openpyxl `Workbook` object directly.
    wb_wrapper.wb
    # Access the wrapped openpyxl `Worksheet` object directly.
    sheet_wrapper.ws

    # Save the changes in the copied .xlsx file, and close.
    wb_wrapper.save_wb()
    wb_wrapper.close_wb()


Classes
-------

.. toctree::
   modules/workbookwrapper
   modules/worksheetwrapper


Indices and tables
==================

* :ref:`genindex`
* :ref:`search`
