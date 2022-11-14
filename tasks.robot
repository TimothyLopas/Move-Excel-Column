*** Settings ***
Documentation       A sample script that moves Column A to the end of the workbook.

Library             RPA.Excel.Files
Library             RPA.Tables


*** Tasks ***
Minimal task
    Creat new workbook with updated content
    Update workbook with new column content
    Update workbook with new column content loop


*** Keywords ***
Creat new workbook with updated content
    Open Workbook    ${CURDIR}${/}Workbook1.xlsx
    ${table}=    Read Worksheet As Table    name=Sheet1    header=True
    Close Workbook
    ${values}=    Get Table Column    ${table}    State
    Open Workbook    ${CURDIR}${/}Workbook2.xlsx
    ${table2}=    Read Worksheet As Table    name=Sheet1    header=True
    Close Workbook
    Add Table Column    ${table2}    name=State Most Sold In    values=${values}
    Create Workbook    ${CURDIR}${/}    fmt=xlsx
    Create Worksheet    New content    content=${table2}    header=True
    Save Workbook    Workbook3.xlsx

Update workbook with new column content
# Use this type of solution only if column names can be compared aross worksheets and each column already has data.
# If used for a new column the data will be appended below the last row in a preceding column
    Open Workbook    ${CURDIR}${/}Workbook1.xlsx
    ${table}=    Read Worksheet As Table    name=Sheet1    header=True
    Close Workbook
    ${values}=    Get Table Column    ${table}    State
    ${column_names}=    Create List    State
    ${values_table}=    Create Table    ${values}    columns=${column_names}
    Open Workbook    ${CURDIR}${/}Workbook2.xlsx
    ${column}=    Set Variable    C
    # Creates the header row so that the Append Rows To Worksheet knows where to place the rest
    Set Cell Value    ${1}    ${column}    State
    Append Rows To Worksheet    ${values_table}    name=Sheet1    header=True    start=${1}
    Save Workbook
    Close Workbook

Update workbook with new column content loop
    Open Workbook    ${CURDIR}${/}Workbook1.xlsx
    ${table}=    Read Worksheet As Table    name=Sheet1    header=True
    Close Workbook
    ${values}=    Get Table Column    ${table}    State
    Open Workbook    ${CURDIR}${/}Workbook4.xlsx
    ${column}=    Set Variable    C
    ${row}=    Set Variable    ${1}
    # Set the header row since the header is not present in the column list of values
    Set Cell Value    ${row}    ${column}    State
    ${row}=    Evaluate    ${row}+1
    FOR    ${value}    IN    @{values}
        Set Cell Value    ${row}    ${column}    ${value}
        ${row}=    Evaluate    ${row}+1
    END
    Save Workbook
    Close Workbook
