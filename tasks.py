from pathlib import Path
from os import environ

from RPA.Excel.Files import Files
from RPA.Tables import Tables

excel = Files()
tables = Tables()

curdir = environ.get("ROBOT_ROOT", "/")

def minimal_task():
    create_new_workbook_with_updated_content()
    update_workbook_with_new_column_content()
    update_workbook_with_new_column_content_loop()


def create_new_workbook_with_updated_content():
    excel.open_workbook(str(Path(curdir, "Workbook1.xlsx")))
    table = excel.read_worksheet_as_table(name="Sheet1", header=True)
    excel.close_workbook()
    values = tables.get_table_column(table, "State")
    excel.open_workbook(str(Path(curdir, "Workbook2.xlsx")))
    table2 = excel.read_worksheet_as_table(name="Sheet1", header=True)
    excel.close_workbook()
    tables.add_table_column(table2, name="State Most Sold In", values=values)
    excel.create_workbook(str(Path(curdir)), fmt="xlsx")
    excel.create_worksheet("New content", content=table2, header=True)
    excel.save_workbook("Workbook3.xlsx")

def update_workbook_with_new_column_content():
# Use this type of solution only if column names can be compared aross worksheets and each column already has data.
# If used for a new column the data will be appended below the last row in a preceding column
    excel.open_workbook(str(Path(curdir, "Workbook1.xlsx")))
    table = excel.read_worksheet_as_table(name="Sheet1", header=True)

    excel.close_workbook()
    values = tables.get_table_column(table, "State")
    column_names = ["State"]
    values_table = tables.create_table(values, columns=column_names)
    excel.open_workbook(str(Path(curdir, "Workbook2.xlsx")))
    column = "C"
    # Creates the header row so that the Append Rows To Worksheet knows where to place the rest
    excel.set_cell_value(1, column, "State")
    excel.append_rows_to_worksheet(values_table, name="Sheet1", header=True, start=1)
    excel.save_workbook()
    excel.close_workbook()


def update_workbook_with_new_column_content_loop():
    excel.open_workbook(str(Path(curdir, "Workbook1.xlsx")))
    table = excel.read_worksheet_as_table(name="Sheet1", header=True)
    excel.close_workbook()
    values = tables.get_table_column(table, "State")
    excel.open_workbook(str(Path(curdir, "Workbook4.xlsx")))
    column = "C"
    row = 1
    # Set the header row since the header is not present in the column list of values
    excel.set_cell_value(row, column, "State")
    row = row + 1
    for value in values:
        excel.set_cell_value(row, column, value)
        row = row + 1
    excel.save_workbook()
    excel.close_workbook()

if __name__ == "__main__":
    minimal_task()
