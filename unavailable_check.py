from openpyxl import load_workbook
from datetime import datetime


def check_people_availability(file_path):
    wb = load_workbook(file_path)
    sheet = wb['USP scheduling']

    unavailable_dict = {}

    # Col and row ranges
    start_row = 5
    end_row = 736
    start_col = 1  # col A
    end_col = 60  # finishing col

    # Searching day by day
    for row in range(start_row, end_row + 1):
        # Reading col A
        date_cell = sheet.cell(row=row, column=1)
        date_value = date_cell.value

        if date_value is None:
            continue

        # Date conversion "month/day/year"
        date_str = date_value.strftime('%m/%d/%Y')

        # Look for unavailability only in selected time range
        month = date_value.month
        current_month = datetime.now().month
        if current_month - 3 <= month <= current_month + 3:
            # Searching in each column for selected row
            for col in range(start_col + 2, end_col + 1):  # + 2, to begin in col C
                person_name = sheet.cell(row=4, column=col).value
                availability = sheet.cell(row=row, column=col).value

                if availability == 'Not available' or availability == 'Holiday':
                    # Add person to list
                    if date_str not in unavailable_dict:
                        unavailable_dict[date_str] = []
                    unavailable_dict[date_str].append(person_name)

    wb.close()

    return unavailable_dict