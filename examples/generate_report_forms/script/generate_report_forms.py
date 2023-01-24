
"""
An example script to take a master spreadsheet that contains data dumped
from expense reports submitted by various teams; and generate new
spreadsheets (one per team) to be filled by somebody who is verifying
each team's documentation.  Specifically:
1) One report per team.
2) Ignore any expenses for items that cost less than $10/item.
3) Add Excel formulas to multiply the price per item by the number of
    items ordered (i.e. multiply columns C and E in each row).

Note: Before calling this script, the master spreadsheet has been
prepared with columns that will eventually be filled in manually by a
person(s) tasked with reviewing the data for each team.
"""

from pathlib import Path
from openpyxl.styles.numbers import BUILTIN_FORMATS
from src import xlsx_copycull

# This format code will be used for cells that we add formulas to.
# Reference number_format codes here -- #44 is Excel's built-in
# 'Accounting' with 2 decimal places:
# https://openpyxl.readthedocs.io/en/stable/_modules/openpyxl/styles/numbers.html
ACCOUNTING_FORMAT = BUILTIN_FORMATS[44]

if __name__ == '__main__':
    # Master spreadsheet location and information.
    master_spreadsheet = Path(r"examples\generate_report_forms\script\original\purchase_data.xlsx")
    sheet_name = 'Accounting'
    header_row = 1

    # Where we'll save the report forms.
    report_directory = Path(r"examples\generate_report_forms\script\reports")

    # The second and third rows contain samples for the reviewers to
    # reference, so we want to keep them in each copy.
    sample_rows = {2, 3}

    # The original spreadsheet has 'Team Code' values from 7 to 23.
    team_codes = range(7, 24)

    # Generate a report form for each team.
    for team_code in team_codes:
        # Filename for each report will encode the team code.
        report_name = f"Team {team_code:02d} Expense Verification Report.xlsx"

        # Copy the master spreadsheet using that filename.
        wb_wrapper = xlsx_copycull.WorkbookWrapper(
            orig_fp=master_spreadsheet,
            copy_fp=report_directory / report_name)
        wb_wrapper.load_wb()

        # We'll keep only items that cost at least $10.00/item, and
        # whose costs were incurred by this team.
        # If this were a SQL query, it might look like:
        #    SELECT *
        #    FROM SomeTable
        #    WHERE price_per_item >= 10 AND team_code = '<this team>';
        select_conditions = {
            'Price Per Item': lambda ppi: ppi >= 10,
            'Team Code': lambda tc: tc == team_code
        }

        # Formula in Column G to multiply each row's C-value by E-value.
        formulas_to_add = {
            'G': lambda row_num: f"=C{row_num}*E{row_num}"
        }
        # Apply 'Accounting' format to each of the resulting formulas
        # in Column G.
        num_formats = {
            'G': ACCOUNTING_FORMAT
        }

        # Stage the 'Accounting' worksheet and rename it for this team.
        ws_wrapper = wb_wrapper.stage_ws(
            ws_name=sheet_name,
            header_row=header_row,
            protected_rows=sample_rows,
            rename_ws=f"{team_code:02d}_expense_verif"
        )

        # Delete unwanted rows
        ws_wrapper.cull(select_conditions=select_conditions, bool_oper='AND')
        # Add formulas, and add appropriate formatting to those same cells
        modified_cells = ws_wrapper.add_formulas(
            formulas_to_add, number_formats=num_formats)

        # If we wanted to modify style on those same cells further, we
        # could iterate over the cells in `modified_cells` (keyed by
        # column letter, i.e. 'G'; whose value is a list of cell names).

        # Save and close.
        wb_wrapper.save_wb()
        wb_wrapper.close_wb()
