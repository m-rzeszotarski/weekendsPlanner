from datetime import datetime


class Day:
    def __init__(self, date):
        self.date = date
        self.day_of_week = self.get_name_of_weekday()
        self.number_of_people = 2
        self.assigned_people = []

    def get_name_of_weekday(self):
        date_obj = datetime.strptime(self.date, '%d/%m/%Y')
        return date_obj.strftime('%A')

    def __str__(self):
        return f"{self.date} ({self.day_of_week})"


class Person:
    def __init__(self, initials, holiday_name, shifts_name):
        self.initials = initials
        self.holiday_name = holiday_name
        self.shifts_name = shifts_name
        self.working_days = []
        self.forced_working_days = []
        self.unavailable = []
        self.holiday = []
        self.working_saturdays = 0
        self.working_sundays = 0
        self.assigned_saturdays = 0
        self.assigned_sundays = 0
        self.exclude = False
