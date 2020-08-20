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
def evaluate_sheet(source_row, col1, col2, col3, col4, step_name):
    # Find the cell and value we want to evaluate
    program_cell = get_cell_by_column_name(source_row, str(col1))
    program_value = program_cell.display_value
    checkbox = False
    type1 = ''
    type2 = ''
    parent = False
    if step_name == 'Subcloning':
        type1 = get_cell_by_column_name(source_row, "Program").value
        parent = get_cell_by_column_name(source_row, "# sorted").value
    if step_name == 'Sequencing':
        checkbox = get_cell_by_column_name(source_row, str(col4)).value
    if step_name == 'CHO Cloning':
        type2 = get_cell_by_column_name(source_row, "Reagent Type").value
    if program_value is not None:
        if (type1 == "anti-drug" and parent) or checkbox or type2 == "Anti-Drug" \
                or step_name == "Protein Biochemistry" or step_name == "Primary Screening" or step_name == "Production":
            start_date_cell = get_cell_by_column_name(source_row, str(col2))
            finish_date_cell = None
            if col3:
                finish_date_cell = get_cell_by_column_name(source_row, str(col3))
            clone_cell = None
            if step_name == "CHO Cloning" or step_name == "Production":
                clone_cell = get_cell_by_column_name(source_row, str(col4))
            # Skip if empty date cell or older than 2 years ago
            if start_date_cell.value and (datetime.now() - datetime.fromisoformat(start_date_cell.value)).days < 730:
                print("Need to add row #" + str(source_row.row_number))
                sdate = ''
                fdate = ''

                # Program Name
                new_cell_1 = smartsheet_client.models.Cell()
                new_cell_1.column_id = result_column_map["Program"]
                if "REGN" in program_value:
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

                # Milestone
                new_cell_2 = smartsheet_client.models.Cell()
                new_cell_2.column_id = result_column_map["Milestone"]
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
                if new_cell_4.value != '':
                    new_cell_6.value = "Completed"
                else:
                    new_cell_6.value = "Cloning In-Progress"
                if step_name == 'CHO Cloning':
                    clone_status = get_cell_by_column_name(source_row, "Status").display_value
                    if clone_status == 'Cancelled':
                        print("This row was cancelled")
                        new_cell_6.value = "Cancelled"
                if step_name == "Protein Biochemistry":
                    purification_status = get_cell_by_column_name(source_row, "Purification Status").display_value
                    if purification_status != '':
                        new_cell_6.value = "Purification " + purification_status
                if step_name == "Production":
                    prod_status = get_cell_by_column_name(source_row, "Production Status").display_value
                    if prod_status != '':
                        new_cell_6.value = "Production " + prod_status

                # Duration
                dur = smartsheet_client.models.Cell()
                dur.column_id = result_column_map["Duration (days)"]
                if not fdate == '' and not sdate == '':
                    duration = fdate - sdate
                    days1 = duration.days
                    dur.value = days1
                else:
                    dur.value = ''

                # REGN# column
                regn = smartsheet_client.models.Cell()
                regn.column_id = result_column_map["REGN#"]
                regn.value = ''
                if step_name == 'CHO Cloning' or step_name == "Protein Biochemistry" or step_name == "Production":
                    temp = get_cell_by_column_name(source_row, "REGN#").value
                    if temp != '':
                        regn.value = temp

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
                new_row.cells.append(regn)

                return new_row

    return None


def get_sheet(sheet_id, rowsToAdd, col1, col2, col3, col4, step_name):
    # Load entire sheet
    sheet = smartsheet_client.Sheets.get_sheet(sheet_id)
    print("Loaded " + str(len(sheet.rows)) + " rows from sheet: " + sheet.name)

    # Build column map for later reference - translates column names to column id
    for c in sheet.columns:
        column_map[c.title] = c.id

    for r in sheet.rows:
        rowToAdd = evaluate_sheet(r, col1, col2, col3, col4, step_name)
        if rowToAdd is not None:
            rowsToAdd.append(rowToAdd)

    return rowsToAdd


def update_sheet(sheet_id):
    # Load entire sheet
    sheet = smartsheet_client.Sheets.get_sheet(sheet_id)
    print("Loaded " + str(len(sheet.rows)) + " rows from sheet: " + sheet.name)

    update = []

    sub_list = []
    for r in sheet.rows:
        mile = r.get_column(result_column_map["Milestone"]).display_value
        if mile == "Subcloning":
            sub_list.append(r)

    for r in sheet.rows:
        mile = r.get_column(result_column_map["Milestone"]).display_value
        if mile == "Primary Screening":
            up  = update_helper(r, sub_list)
            update.append(up)

    return update


def update_helper(source_row, sub_list):
    # Finish date
    new_cell = smartsheet.models.Cell()
    new_cell.column_id = result_column_map["End"]
    new_cell.value = ""

    # Duration
    dur = smartsheet_client.models.Cell()
    dur.column_id = result_column_map["Duration (days)"]
    dur.value = ""

    # Status
    stat = smartsheet.models.Cell()
    stat.column_id = result_column_map["Status"]
    stat.value = "Primary Screening In-Progress"

    for sub in sub_list:
        primary_prog = source_row.get_column(result_column_map["Program"]).display_value
        if primary_prog == sub.get_column(result_column_map["Program"]).value:
            new_cell.value = sub.get_column(result_column_map["Start"]).value
            duration = (datetime.fromisoformat(sub.get_column(result_column_map["End"]).value) - datetime.fromisoformat(sub.get_column(result_column_map["Start"]).value)).days
            dur.value = duration
            stat.value = "Completed"

    new_row = smartsheet_client.models.Row()
    new_row.id = source_row.id
    new_row.cells.append(new_cell)
    new_row.cells.append(dur)
    new_row.cells.append(stat)
    return new_row


def delete_excess(sheet_id):
    sheet = smartsheet_client.Sheets.get_sheet(sheet_id)
    print("Loaded " + str(len(sheet.rows)) + " rows from sheet: " + sheet.name)
    d_rows = []
    prev_program = ''
    for r in sheet.rows:
        flag, prev_program = delete_helper(r, "Milestone", prev_program)
        if flag:
            d_rows.append(r.id)
    return d_rows


def delete_helper(source_row, step_name, prev_program):
    flag = False
    next_program = source_row.get_column(result_column_map["Program"]).display_value
    curr_milestone = source_row.get_column(result_column_map[step_name]).display_value
    if prev_program == next_program and curr_milestone == "Sequencing":
        flag = True
    return flag, next_program


def clear_rows(sheet_id, rows_to_delete):
    num = len(rows_to_delete)
    step = 0.25
    for i in range(4):
        mu1 = step * i
        mu2 = step * (i + 1)
        temp = rows_to_delete[int(mu1 * num):int(mu2 * num)]
        smartsheet_client.Sheets.delete_rows(sheet_id, temp)
    print("Cleared " + str(len(rows_to_delete)) + " rows")


print("Starting ...")

# Initialize client. Uses the API token in the environment variable "SMARTSHEET_ACCESS_TOKEN"
smartsheet_client = smartsheet.Smartsheet("3dwmt4ylby5aomwb3irzud3u7w")
# Make sure we don't miss any error
smartsheet_client.errors_as_exceptions(True)

# Log all calls
logging.basicConfig(filename='rwsheet.log', level=logging.INFO)

# IDs of Smartsheets to import
result_id = 6450008561084292
primary_screen_id = 6583769915254660
subcloning_id = 8954984113956740
sequencing_id = 3900772095158148
cho_cloning_id = 6665824795682692
pb_id = 5973156864255876
production_id = 8867370774095748
ad_tracker_id = 6250422773016452

result_sheet = smartsheet_client.Sheets.get_sheet(result_id)

# Clear the master sheet
rowsToDelete = []
for row in result_sheet.rows:
    rowsToDelete.append(row.id)
if len(rowsToDelete) > 2:
    clear_rows(result_id, rowsToDelete)

for column in result_sheet.columns:
    result_column_map[column.title] = column.id

# Accumulate rows needing update here
rowsToAdd = []
rowsToAdd = get_sheet(primary_screen_id, rowsToAdd, "Target", "Primary Screen Date", "", "", "Primary Screening")
rowsToAdd = get_sheet(subcloning_id, rowsToAdd, "Target / Hybridoma ID", "Sort Date", "Finish", "", "Subcloning")
rowsToAdd = get_sheet(sequencing_id, rowsToAdd, "Target", "NGS Run Date", "", "Anti-Drug", "Sequencing")
rowsToAdd = get_sheet(cho_cloning_id, rowsToAdd, "Target", "Assigned Date", "Cloning Complete", "AbPID/REGN#", "CHO Cloning")
rowsToAdd = get_sheet(pb_id, rowsToAdd, "Program", "Purification Start Date", "Purification Completion Date", "", "Protein Biochemistry")
rowsToAdd = get_sheet(production_id, rowsToAdd, "Target", "Cloning Complete", "Production Complete", "AbPID/REGN#", "Production")
# rowsToAdd = get_sheet(ad_tracker_id, rowsToAdd, "Target", "Cloning Complete", "Production Complete", "AbPID/REGN#", "ADG Request")

# Finally, write updated cells back to Smartsheet
if rowsToAdd:
    print("Writing " + str(len(rowsToAdd)) + " rows back to sheet id " + str(result_id))
    response = smartsheet_client.Sheets.add_rows(result_id, rowsToAdd)
else:
    print("No rows to add!")

# Delete excess sequencing rows
d = delete_excess(result_id)
clear_rows(result_id, d)

# Update Primary Screening end date:
rowsToUpdate = update_sheet(result_id)
updated_row = smartsheet_client.Sheets.update_rows(result_id, rowsToUpdate)
print("Updated " + str(len(rowsToUpdate)) + " rows for Primary Screening dates")

print("Done")