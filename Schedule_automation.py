from ctypes import wstring_at
from tkinter.messagebox import showerror

import openpyxl
from openpyxl import Workbook, load_workbook
from openpyxl.utils.cell import coordinate_from_string
from openpyxl.styles import Font
from datetime import time, datetime
from pathlib import Path
import pyexcel_io.writers
import pyexcel_io.readers
import pyexcel as p
import os

from xlrd.formula import sheetrange


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

def get_travel_time(source_file, sheet_index, target):

    wb_source = openpyxl.load_workbook(source_file)

    if sheet_index >=  len(wb_source.sheetnames) or sheet_index < 0:
        print(f"Error: Sheet  index'{sheet_index}' is out of range.")
        return[]

    target_sheet_name = wb_source.sheetnames[sheet_index]
    ws_source = wb_source[target_sheet_name]

    search_term = str(target).lower()
    found_instances = []

    for row in ws_source.iter_rows():
        for cell in row:
            if cell.value and search_term in str(cell.value).lower():
                found_instances.append(
                    {
                        "coordinate": cell.offset(row=0, column=1).coordinate
                    }
                )
    return found_instances


def get_header_coordinates(source_file, sheet_index, target_header):

    wb_source = openpyxl.load_workbook(source_file, read_only=True)

    if sheet_index >=  len(wb_source.sheetnames) or sheet_index < 0:
        print(f"Error: Sheet  index'{sheet_index}' is out of range.")
        return[]

    target_sheet_name = wb_source.sheetnames[sheet_index]
    ws_source = wb_source[target_sheet_name]

    print(f"Scanning Sheet Index {sheet_index} (Named: '{target_sheet_name}')...")

    search_term = str(target_header).lower()
    found_instances = []

    for row in ws_source.iter_rows():
        for cell in row:
            if cell.value and search_term in str(cell.value).lower():
                found_instances.append(
                    {
                        "coordinate": cell.coordinate,
                        "column_letter": cell.column_letter,
                        "data row": cell.column_letter + str(cell.row + 1),
                        "column_index": cell.column,
                    }
                )
    return found_instances

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
    
def scheme_and_trip_duration ( #todo Remove the fwd_orientation part (unneeded) and make it apply trip duration seperately for fwd and bwd
        source_wb_path,
        out_wb_path,
        trip_duration,
        fwd_orientation
):
    wb_source = load_workbook(source_wb_path)
    wb_out = load_workbook(out_wb_path)

    if fwd_orientation == True:
        forward = "Forward"
        backward = "Backward"
    else:
        forward = "Backward"
        backward = "Forward"

    shift_change_time = datetime.strptime("14:00", "%H:%M").time()

    for idx, sheet in enumerate(wb_out.worksheets):

        max_row = 0
        for row in range(1, sheet.max_row + 1):
            if sheet.cell(row=row, column=1).value or sheet.cell(row=row, column=4).value:
                max_row = row

        for row in range(2, max_row + 1):
            time_str = sheet.cell(row=row, column=4).value

            shift = ""

            if isinstance(time_str, str):
                try:
                    time_obj = datetime.strptime(time_str.strip(), "%H:%M").time()
                    shift = "Morning shift" if time_obj < shift_change_time else "Night shift"
                except ValueError:
                    shift = "ERROR!"

            sheet.cell(row=row, column=2).value = shift
            sheet.cell(row=row, column=5).value = trip_duration
            sheet.cell(row=row, column=6).value = trip_duration
            sheet.cell(row=row, column=3).value = forward if idx == 0 else backward

    wb_out.save(out_wb_path)

def run_excel_stuff(
        input_path_from_user
):

    # TODO WIDEN THE NET HERE TO CATCH INCONSISTENT FILE NAMES
    # List of routes - for identifying what route we're working on - some of these arent what we use in the output file
    list_of_routes = ["EXP-01","SR-02", "DR-3A", "DR-3B", "DR-4B", "DR-05", "DR-06", "DR-07", "SR-08", "EXP-09", "EXP-10",
                      "DR-11", "EXP-12", "DR-13", "DR-14", "DR-14A", "XER-15", "EXP-16"]
    # Pretty Bodgy, might be a better way to do this
    string_map_routes = {"EXP-01": "ER-01","SR-02": "SR-02", "DR-3A": "DR-03A", "DR-3B": "DR-03A", "DR-4B": "DR-04B", "DR-05": "DR-05", "DR-06": "DR-06",
                        "DR-07" : "DR-07", "SR-08": "SR-08", "EXP-09": "ER-09", "EXP-10": "ER-10", "DR-11": "DR-11", "EXP-12": "ER-12", "DR-13": "DR-13",
                         "DR-14": "DR-014", "DR-14A": "DR-014A", "XER-15": "XER-15", "EXP-16": "ER-16" }

    normal_routes = ["ER-01", "SR-02", "DR-03A", "DR-03B", "DR-05", "DR-06", "DR-07", "SR-08", "ER-09",
                     "DR-11", "DR-13", "DR-014", "DR-014A", "XER-15", "ER-16"]
    queer_routes = ["ER-10", "ER-12", "DR-04B"]
    normal_set = set(normal_routes)
    queer_set = set(queer_routes)

    # Preparing input worksheet -
    path = input_path_from_user # here is the input file path - hook up to GUI
    just_the_name = os.path.basename(path)
    identifier = next(
        (x for x in list_of_routes if x in Path(path).stem),
        None
    )

    if identifier is None:
        raise ValueError(f"No valid route name found in filename - please check input file")

    print(identifier)
    print(string_map_routes[identifier])

    wb_in = openpyxl.load_workbook(path)
    wb_in2 = openpyxl.load_workbook(path, data_only=True) # why is only this one got data_only=True ?
    ws_stats = wb_in2.worksheets[0]
    ws_details = wb_in2.worksheets[1]

    # Preparing output worksheet
    wb_out = Workbook()

    # hooks for determining route - also take into account some routes have the order backwards
    # - would need to flip the actual data instead and not the order here....
    wb_out_forward_sheet = string_map_routes[identifier] + "(Forward)"
    wb_out_backward_sheet = string_map_routes[identifier] + "(Backward)"

    wb_out_forward = wb_out.create_sheet(wb_out_forward_sheet)
    wb_out_backward = wb_out.create_sheet(wb_out_backward_sheet)
    if 'Sheet' in wb_out.sheetnames:
        del wb_out['Sheet']

    # Write Headers (IMP for schedule)
    schedule_header = ["Shift Number","Shift Type", "Scheme", "Departure Time", "Min Trip Duration", "Max Trip Duration"]
    write_header(wb_out_forward, schedule_header)
    write_header(wb_out_backward, schedule_header)
    # temporary destination output - BODGE
    dest_path = "output_temp.xlsx"
    wb_out.save(dest_path)

    # Bodge - inputs sheets are inconsistent

    column_locations_bus_no = get_header_coordinates(
        source_file=path,
        sheet_index=1,
        target_header= "Bus No"
    )

    print(f"Found {len(column_locations_bus_no)} columns with that name:\n")

    for index, item in enumerate(column_locations_bus_no, 1):
        print(f"Instance {index}:")
        print(f"   -> Data row: {item['data row']}")
        print("-" * 30)


    column_locations_trip_start = get_header_coordinates(
        source_file=path,
        sheet_index=1,
        target_header= "Trip Start Time"
    )

    print(f"Found {len(column_locations_trip_start)} columns with that name:\n")

    for index, item in enumerate(column_locations_trip_start, 1):
        print(f"Instance {index}:")
        print(f"   -> Data row: {item['data row']}")
        print("-" * 30)


    # Copying the columns to their proper places, checking if the route is normal or queer
    # TODO Handle DR014A and DR014 - ADD TO QUEER ROUTES - THERES A HIDDEN TABLE THAT MUST BE IGNORED
    match string_map_routes[identifier]:
        case x if x in normal_set:
            copy_column_as_text_between_workbooks(
                path,
                1,
                column_locations_bus_no[0]["data row"],
                dest_path,
                0,
                "A2",
                False
            )

            copy_column_as_text_between_workbooks(
                path,
                1,
                column_locations_trip_start[0]["data row"],
                dest_path,
                0,
                "D2",
                True
            )

            copy_column_as_text_between_workbooks(
                path,
                1,
                column_locations_bus_no[1]["data row"],
                dest_path,
                1,
                "A2",
                False
            )

            copy_column_as_text_between_workbooks(
                path,
                1,
                column_locations_trip_start[1]["data row"],
                dest_path,
                1,
                "D2",
                True
            )

        case x if x in queer_routes:
            print("how queer...")
            copy_column_as_text_between_workbooks(
                path,
                1,
                column_locations_bus_no[0]["data row"],
                dest_path,
                1,
                "A2",
                False
            )

            copy_column_as_text_between_workbooks(
                path,
                1,
                column_locations_trip_start[0]["data row"],
                dest_path,
                1,
                "D2",
                True
            )

            copy_column_as_text_between_workbooks(
                path,
                1,
                column_locations_bus_no[1]["data row"],
                dest_path,
                0,
                "A2",
                False
            )

            copy_column_as_text_between_workbooks(
                path,
                1,
                column_locations_trip_start[1]["data row"],
                dest_path,
                0,
                "D2",
                True
            )

        case _:
            print("error")

    travel_time = get_travel_time(
        path,
        1,
        "Travel Time"
    )

    print(f"Here is the travel time stuff {travel_time[0]}, {travel_time[1]}")

    # todo TRIP DURATION AUTOMATIC ASSIGNMENT (PROBABLY ANOTHER LIST...) - No! List bad! take it from the input sheet!!!
    # todo need to seperate this to apply the trip duration differently to fwd and bwd, also the value for SR02 will be "01:05" (hours: minutes) -> will need to convert this to "65" (minutes)
    # todo for other routes it will be like this e.g "24:00" (minutes: seconds) [-_-] -> "24" (minutes)
    scheme_and_trip_duration(
        path,
        dest_path,
        "20",
        True
    )

    trim_empty_rows_and_columns(dest_path)

    # Converting the xlsx file (cant use) to xls file and deleting the old file
    convert_xlsx_to_xls(dest_path,f"{just_the_name} - DONE.xls")
    os.remove(dest_path)



if __name__ == "__main__":
    run_excel_stuff()