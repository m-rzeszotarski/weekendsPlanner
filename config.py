# Here you can add initials of people who will be excluded as default, so you don't have to mark them each time you run
# software. Stick to the python list format so for example ['AZY', 'MIRZ', 'WIA']
DEFAULT_EXCLUDED_PEOPLE = ['AZY']

# Holiday file name - if found in main directory it will be loaded automatically
HOLIDAY_FILE_NAME = "02.2022-02.2023 Holiday calendar.xlsx"
# Shifts file name - if found in main directory it will be loaded automatically
SHIFTS_FILE_NAME = "2022-2024 USP TR&D shifts.xlsx"
# Names/initials file name - name of csv where info about employees names in holiday and shifts files are stored
NAMES_INITIALS_FILE_NAME = "names_initials.csv"

# Configuration of Holiday file reading

# Name of the sheet with holiday calendar
HOLIDAY_SHEET_NAME = 'USP scheduling'
# Row and columns ranges for searching
HOLIDAY_START_ROW = 5
HOLIDAY_END_ROW = 736
# Column with dates
HOLIDAY_START_COL = 1
# Last column to check
HOLIDAY_END_COL = 60
# First column with employee name
HOLIDAY_NAMES_START_COL = 3
# A list of words that, when read by the program, will consider a person as unavailable on a given day
UNAVAILABLE_WORDS_LIST = ['Holiday', 'Not available', 'Official Holiday']
# Range of months +/- in which the program looks for people unavailability
HOLIDAY_MONTHS_SEARCH_RANGE = 3

# Configuration of Shifts file reading

# A year in which program counts working saturdays and sundays. Value must enter in quotation marks (string).
YEAR = '2024'


MAIN_WINDOW_RESOLUTION = "600x800"
EXCLUDE_WINDOW_RESOLUTION = "300x400"
