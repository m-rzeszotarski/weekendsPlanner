import unittest
from holiday_calendar_read import check_people_fields
from shifts_file_read import count_weekend_days
from main import read_initials_file
import config
from unittest.mock import patch
from tkinter import messagebox


class TestWeekendsPlanner(unittest.TestCase):

    def setUp(self):
        self.people_in_team = read_initials_file()

    @patch('tkinter.messagebox.showinfo')  # Mocking showinfo to prevent opening the dialog
    def test_read_holidays(self, mock_messagebox):
        holidays = check_people_fields("02.2022-02.2023 Holiday calendar.xlsx", config.UNAVAILABLE_WORDS_LIST)
        self.assertIsInstance(holidays, dict)
        self.assertTrue(len(holidays) > 0)  # Check if holidays were loaded

    @patch('tkinter.messagebox.showinfo')  # Mocking showinfo again in the second test
    def test_read_shifts(self, mock_messagebox):
        shifts = count_weekend_days("2022-2024 USP TR&D shifts.xlsx", self.people_in_team)
        self.assertIsInstance(shifts, dict)
        self.assertTrue(len(shifts) > 0)  # Check if the shift data was loaded


if __name__ == "__main__":
    unittest.main()