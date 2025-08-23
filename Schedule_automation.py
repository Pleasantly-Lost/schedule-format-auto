import openpyxl
from openpyxl import Workbook, load_workbook
from openpyxl.utils.cell import coordinate_from_string
from openpyxl.styles import Font
from datetime import time
import pyexcel as p
import os

def copy_column_as_text_between_workbooks(
    source_wb_path,
    source_sheet_index,
    source_cell_ref,
    target_wb_path,
    target_sheet_index,
    target_cell_ref,
    convert_time_to_hhmm=False
):
    # Load both workbooks
    source_wb = load_workbook(source_wb_path, data_only=True)  # data_only=True avoids formulas
    target_wb = load_workbook(target_wb_path)

    # Get worksheets
    ws_source = source_wb.worksheets[source_sheet_index]
    ws_target = target_wb.worksheets[target_sheet_index]

    # Parse source and target cell references
    source_col_letter, source_start_row = coordinate_from_string(source_cell_ref)
    target_col_letter, target_start_row = coordinate_from_string(target_cell_ref)

    # Step 1: Find the last non-empty cell in the source column
    current_row = source_start_row
    while True:
        if ws_source[f"{source_col_letter}{current_row}"].value is None:
            break
        current_row += 1
    end_row = current_row - 1

    # Step 2: Copy values to target, formatting as Text
    for i, row in enumerate(ws_source[f"{source_col_letter}{source_start_row}:{source_col_letter}{end_row}"]):
        source_cell = row[0]
        value = source_cell.value
        target_row = target_start_row + i
        target_cell = ws_target[f"{target_col_letter}{target_row}"]

        if convert_time_to_hhmm and isinstance(value, time):
            # Format time as HH:MM and save as string
            time_str = value.strftime('%H:%M')
            target_cell.value = time_str
            target_cell.number_format = '@'
        else:
            # Save other values as string with text format
            target_cell.value = str(value) if value is not None else ""
            target_cell.number_format = '@'

    # Save the target workbook only
    target_wb.save(target_wb_path)


def trim_empty_rows_and_columns(filename):
    wb = load_workbook(filename)

    for ws in wb.worksheets:
        # Find last row and column with data
        max_row = 0
        max_col = 0
        for row in ws.iter_rows():
            for cell in row:
                if cell.value is not None:
                    max_row = max(max_row, cell.row)
                    max_col = max(max_col, cell.column)

        # Delete rows below the last data row
        if max_row < ws.max_row:
            ws.delete_rows(max_row + 1, ws.max_row - max_row)

        # Delete columns after the last data column
        if max_col < ws.max_column:
            ws.delete_cols(max_col + 1, ws.max_column - max_col)

    wb.save(filename)

def write_header(sheet, header):
    for col_num, header_value in enumerate(header, start=1):
        cell = sheet.cell(row=1, column=col_num, value=header_value)
        cell.font = Font(bold=True)

def convert_xlsx_to_xls(xlsx_file, xls_file):
    # Read the .xlsx file
    sheet = p.get_book(file_name=xlsx_file)

    # Save as .xls
    sheet.save_as(xls_file)

def run_excel_stuff():####################### hook up path and dest_path as inputs to this function


    # Preparing output worksheet
    wb_out = Workbook()
    wb_out_forward = wb_out.create_sheet("Forward")
    wb_out_backward = wb_out.create_sheet("Backward")
    if 'Sheet' in wb_out.sheetnames:
        del wb_out['Sheet']
    # Write Headers (IMP for schedule)
    schedule_header = ["Shift Number","Shift Type", "Scheme", "Departure Time", "Min Trip Duration", "Max Trip Duration"]
    write_header(wb_out_forward, schedule_header)
    write_header(wb_out_backward, schedule_header)

    dest_path = "output_test.xlsx"
    wb_out.save(dest_path)


    # Preparing input worksheet
    path = "3. DR-3A Schedule Tue, Wed & Fri August 1st, 2025.xlsx" # here is the input file path
    wb_in = openpyxl.load_workbook(path)
    wb_in2 = openpyxl.load_workbook(path, data_only=True)

    ws_stats = wb_in2.worksheets[0]
    ws_details = wb_in2.worksheets[1]

    forward_total = ws_stats['F28']
    backward_total = ws_stats['F38']


    # what was this for again?
    # start_row = 12
    # column_letter = 'B'
    #
    # end_row = start_row
    # while True:
    #     cell_value = ws_details[f"{column_letter}{end_row}"].value
    #     if cell_value is None:
    #         break
    #     end_row += 1
    #
    # cell_range = ws_details[f"{column_letter}{start_row}:{column_letter}{end_row - 1}"]
    #
    # target_start_row = 2
    # target_col_letter = 'A'
    #
    # for i, row in enumerate(cell_range):
    #     value = row[0].value
    #     target_cell = ws_out[f"{target_col_letter}{target_start_row + i}"]
    #     target_cell.value = str(value) if value is not None else ""
    #     target_cell.number_format = '@'


    ################# OUTPUT SHEET HERE IS WRONG IN THIS FUNCTION, IT MAKES A NEW OUTPUT SHEET BY ITSELF WHEN IT SHOULD ONLY MODIFY THE ONE WE MADE -- FIXED

    copy_column_as_text_between_workbooks(
        path,
        1,
        "B12",
        dest_path,
        0,
        "A2",
        False
    )

    copy_column_as_text_between_workbooks(
        path,
        1,
        "D12",
        dest_path,
        0,
        "D2",
        True
    )

    copy_column_as_text_between_workbooks(
        path,
        1,
        "I12",
        dest_path,
        1,
        "A2",
        False
    )

    copy_column_as_text_between_workbooks(
        path,
        1,
        "K12",
        dest_path,
        1,
        "D2",
        True
    )

    trim_empty_rows_and_columns(dest_path)
    convert_xlsx_to_xls(dest_path,"Test format test.xls")
    os.remove(dest_path)




