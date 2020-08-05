# Install the smartsheet sdk with the command: pip install smartsheet-python-sdk
import smartsheet
import logging
import os
from datetime import datetime, timedelta

_dir = os.path.dirname(os.path.abspath(__file__))

# The API identifies columns by Id, but it's more convenient to refer to column names. Store a map here
column_map = {}
result_column_map = {}

# Helper function to find cell in a row
def get_cell_by_column_name(row, column_name):
    column_id = column_map[column_name]
    return row.get_column(column_id)


# TODO: Replace the body of this function with your code
# This *example* looks for rows with a "Status" column marked "Complete" and sets the "Remaining" column to zero
# Return a new Row with updated cell values, else None to leave unchanged
def evaluate_sheet(source_row, col1, col2, col3, col4, step_name, prev_program):
    # Find the cell and value we want to evaluate
    program_cell = get_cell_by_column_name(source_row, str(col1))
    program_value = program_cell.display_value
    checkbox = False
    if step_name == 'Sequencing':
        checkbox = get_cell_by_column_name(source_row, str(col4))
    if (program_value is not None and 'REGN' in program_value) or (program_value is not None and checkbox):
        if step_name == 'Sequencing' and program_value == prev_program:
            return None, None
        start_date_cell = get_cell_by_column_name(source_row, str(col2))
        finish_date_cell = None
        if col3:
            finish_date_cell = get_cell_by_column_name(source_row, str(col3))
        clone_cell = None
        if step_name == "CHO Cloning":
            clone_cell = get_cell_by_column_name(source_row, str(col4))
        if start_date_cell.value:  # Skip if empty date cell
            print("Need to add row #" + str(source_row.row_number))
            sdate = ''
            fdate = ''
            # Program Name
            new_cell_1 = smartsheet_client.models.Cell()
            new_cell_1.column_id = result_column_map["Program"]
            if not checkbox:
                i = program_value.find('REGN')
                k = program_value.find(' ')
                j = program_value.find('(')
                if k != -1 and k > i:
                    new_cell_1.value = program_value[i:k]
                elif j != -1 and j > i:
                    new_cell_1.value = program_value[i:j]
                else:
                    new_cell_1.value = program_value[i:]
            else:
                new_cell_1.value = program_value

            # Step Name
            new_cell_2 = smartsheet_client.models.Cell()
            new_cell_2.column_id = result_column_map["Step"]
            new_cell_2.value = step_name

            # Start date
            new_cell_3 = smartsheet_client.models.Cell()
            new_cell_3.column_id = result_column_map["Start"]
            if start_date_cell is not None and start_date_cell.value is not None:
                new_cell_3.value = start_date_cell.value
                sdate = datetime.fromisoformat(start_date_cell.value)
            else:
                new_cell_3.value = ''

            # Finish date
            new_cell_4 = smartsheet_client.models.Cell()
            new_cell_4.column_id = result_column_map["End"]
            if finish_date_cell is not None and finish_date_cell.value is not None:
                new_cell_4.value = finish_date_cell.value
                fdate = datetime.fromisoformat(finish_date_cell.value)
            else:
                if step_name == "Sequencing" and sdate != '':
                    new_cell_4.value = (sdate + timedelta(days=7)).isoformat()
                    fdate = sdate + timedelta(days=7)
                elif step_name == "Subcloning" and sdate != '':
                    new_cell_4.value = (sdate + timedelta(days=21)).isoformat()
                    fdate = sdate + timedelta(days=21)
                else:
                    new_cell_4.value = ''

            # Clone name
            new_cell_5 = smartsheet_client.models.Cell()
            new_cell_5.column_id = result_column_map["Clone"]
            if clone_cell:
                new_cell_5.value = clone_cell.value
            else:
                new_cell_5.value = ''

            # Status
            new_cell_6 = smartsheet_client.models.Cell()
            new_cell_6.column_id = result_column_map["Status"]
            if new_cell_4.value != '' or step_name == 'Primary Screening':
                new_cell_6.value = "Complete"
            else:
                new_cell_6.value = "In-Progress"

            # Duration
            dur = smartsheet_client.models.Cell()
            dur.column_id = result_column_map["Duration (days)"]
            if not fdate == '' and not sdate == '':
                duration = fdate - sdate
                days1 = duration.days
                dur.value = days1
            else:
                dur.value = ''

            # Build the row to update
            new_row = smartsheet_client.models.Row()
            new_row.to_bottom = True
            new_row.cells.append(new_cell_1)
            new_row.cells.append(new_cell_2)
            new_row.cells.append(new_cell_3)
            new_row.cells.append(new_cell_4)
            new_row.cells.append(new_cell_5)
            new_row.cells.append(new_cell_6)
            new_row.cells.append(dur)

            return new_row, program_value

    return None, None


def get_sheet(sheet_id, rowsToAdd, col1, col2, col3, col4, step_name):
    # Load entire sheet
    sheet = smartsheet_client.Sheets.get_sheet(sheet_id)
    print("Loaded " + str(len(sheet.rows)) + " rows from sheet: " + sheet.name)

    # Build column map for later reference - translates column names to column id
    for c in sheet.columns:
        column_map[c.title] = c.id

    prev_program = ''
    for row in sheet.rows:
        rowToAdd, prev_program = evaluate_sheet(row, col1, col2, col3, col4, step_name, prev_program)
        if rowToAdd is not None:
            rowsToAdd.append(rowToAdd)
            # if step_name == 'Sequencing':
            #     break

    return rowsToAdd


print("Starting ...")

# Initialize client. Uses the API token in the environment variable "SMARTSHEET_ACCESS_TOKEN"
smartsheet_client = smartsheet.Smartsheet("3dwmt4ylby5aomwb3irzud3u7w")
# Make sure we don't miss any error
smartsheet_client.errors_as_exceptions(True)

# Log all calls
logging.basicConfig(filename='rwsheet.log', level=logging.INFO)

# Import the sheet
# result = smart.Sheets.import_xlsx_sheet(_dir + '/Sample Sheet.xlsx', header_row_index=0)
result = 6450008561084292
primary_screen_id = 6583769915254660
subcloning_id = 8954984113956740
sequencing_id = 3900772095158148
cho_cloning_id = 6665824795682692

result_sheet = smartsheet_client.Sheets.get_sheet(result)

# Clear the master sheet
rowsToDelete = []
for row in result_sheet.rows:
    rowsToDelete.append(row.id)
if len(rowsToDelete) > 2:
    response = smartsheet_client.Sheets.delete_rows(result, rowsToDelete)

for column in result_sheet.columns:
    result_column_map[column.title] = column.id

# Accumulate rows needing update here
rowsToAdd = []
rowsToAdd = get_sheet(primary_screen_id, rowsToAdd, "Target", "Primary Screen Date", "", "", "Primary Screening")
rowsToAdd = get_sheet(subcloning_id, rowsToAdd, "Target / Hybridoma ID", "Sort Date", "Finish", "", "Subcloning")
rowsToAdd = get_sheet(sequencing_id, rowsToAdd, "Target", "NGS Run Date", "", "Anti-Drug", "Sequencing")
rowsToAdd = get_sheet(cho_cloning_id, rowsToAdd, "Target", "Assigned Date", "Cloning Complete", "AbPID/REGN#", "CHO Cloning")

# Finally, write updated cells back to Smartsheet
if rowsToAdd:
    print("Writing " + str(len(rowsToAdd)) + " rows back to sheet id " + str(result))
    response = smartsheet_client.Sheets.add_rows(result, rowsToAdd)
else:
    print("No updates required")

print("Done")
