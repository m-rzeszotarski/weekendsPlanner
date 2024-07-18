# Here you can add initials of people who will be excluded as default, so you don't have to mark them each time you run
# software. Stick to the python list format so for example ['AZY', 'MIRZ', 'WIA']
DEFAULT_EXCLUDED_PEOPLE = ['AZY']

# Configuration of Holiday file reading

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
UNAVAILABLE_WORDS_LIST = ['Holiday', 'Not available']
# Range of months +/- in which the program looks for people unavailability
HOLIDAY_MONTHS_SEARCH_RANGE = 3

# Configuration of Shifts file reading

# A year in which program counts working saturdays and sundays. Value must enter in quotation marks (string).
YEAR = '2024'
