import random
from datetime import datetime, timedelta
from tkinter import messagebox
from models import Person
from write_shifts import generate_shifts_sheet
import config


def assign_forced_people(day, people):
    assigned_people = []
    for person in people:
        if day.date in person.forced_working_days:
            assigned_people.append(person)
            person.working_days.append(day.date)
            if day.day_of_week == 'Saturday':
                person.working_saturdays += 1
                person.assigned_saturdays += 1
            elif day.day_of_week == 'Sunday':
                person.working_sundays += 1
                person.assigned_sundays += 1
    return assigned_people


def assign_people(people_list, day, num_needed, assigned_people):
    random.shuffle(people_list)
    for person in people_list:
        if num_needed <= 0:
            break
        if not any(
                day.date == working_day or
                (datetime.strptime(day.date, '%d/%m/%Y') - timedelta(days=1)).strftime('%d/%m/%Y')
                == working_day or
                (datetime.strptime(day.date, '%d/%m/%Y') + timedelta(days=1)).strftime('%d/%m/%Y')
                == working_day
                for working_day in person.working_days
        ):
            assigned_people.append(person)
            person.working_days.append(day.date)
            if day.day_of_week == 'Saturday':
                person.working_saturdays += 1
                person.assigned_saturdays += 1
            elif day.day_of_week == 'Sunday':
                person.working_sundays += 1
                person.assigned_sundays += 1
            num_needed -= 1


def assign_people_to_days(days, people):
    check_dates(days)
    for day in days:

        assigned_people = assign_forced_people(day, people)

        available_people = [
            person for person in people
            if not person.exclude and day.date not in person.unavailable
        ]

        available_people_with_correct_ratio = [
            person for person in available_people
            if (day.day_of_week == 'Saturday' and person.working_saturdays < person.working_sundays) or
               (day.day_of_week == 'Sunday' and person.working_sundays < person.working_saturdays)
        ]

        available_people_with_correct_ratio_wd1 = [
            person for person in available_people_with_correct_ratio
            if person.assigned_sundays + person.assigned_saturdays == 1

        ]

        available_people_with_correct_ratio_wd2 = [
            person for person in available_people_with_correct_ratio
            if person.assigned_sundays + person.assigned_saturdays == 2

        ]

        available_people_with_correct_ratio_nw = [
            person for person in available_people_with_correct_ratio
            if not person.working_days
        ]

        available_people_wd1 = [
            person for person in available_people
            if person.assigned_sundays + person.assigned_saturdays == 1
        ]

        available_people_wd2 = [
            person for person in available_people
            if person.assigned_sundays + person.assigned_saturdays == 2
        ]

        available_people_nw = [
            person for person in available_people
            if not person.working_days
        ]

        num_needed = day.number_of_people - len(assigned_people)
        assign_people(available_people_with_correct_ratio_nw, day, num_needed, assigned_people)
        num_needed = day.number_of_people - len(assigned_people)
        if num_needed > 0:
            assign_people(available_people_nw, day, num_needed, assigned_people)
        num_needed = day.number_of_people - len(assigned_people)
        if num_needed > 0:
            assign_people(available_people_with_correct_ratio_wd1, day, num_needed, assigned_people)
        num_needed = day.number_of_people - len(assigned_people)
        if num_needed > 0:
            assign_people(available_people_wd1, day, num_needed, assigned_people)
        num_needed = day.number_of_people - len(assigned_people)
        if num_needed > 0:
            assign_people(available_people_with_correct_ratio_wd2, day, num_needed, assigned_people)
        num_needed = day.number_of_people - len(assigned_people)
        if num_needed > 0:
            assign_people(available_people_wd2, day, num_needed, assigned_people)
        num_needed = day.number_of_people - len(assigned_people)
        if num_needed > 0:
            assign_people(available_people_with_correct_ratio, day, num_needed, assigned_people)
        num_needed = day.number_of_people - len(assigned_people)
        if num_needed > 0:
            assign_people(available_people, day, num_needed, assigned_people)

        while len(assigned_people) < day.number_of_people:
            result = messagebox.askyesno("Warning", f"No available people at {day.date}. Do you want to assign random "
                                                    f"person?")
            if result:
                random_person = random.choice(people)
                assigned_people.append(random_person)
                random_person.working_days.append(day.date)
                if day.day_of_week == 'Saturday':
                    random_person.working_saturdays += 1
                    random_person.assigned_saturdays += 1
                elif day.day_of_week == 'Sunday':
                    random_person.working_sundays += 1
                    random_person.assigned_sundays += 1
            else:
                # If no person assigned, generate X in assignment list
                x_person = Person("X", "", "")
                assigned_people.append(x_person)

        day.assigned_people = [person.initials for person in assigned_people]

    summary = "\n".join([f"{day.date} - {', '.join(day.assigned_people)}" for day in days])
    result = messagebox.askyesno("Assignments", f"{summary}\n\nWould you like to generate Excel file with assignments?")
    if result:
        generate_shifts_sheet(days, people)
    reset_assignments(people, days)


# Check if selected day in the calendar was checked by script for people unavailability
def check_dates(days):
    now = datetime.now()
    lower_bound = now - timedelta(days=config.HOLIDAY_MONTHS_SEARCH_RANGE * 30)
    upper_bound = now + timedelta(days=config.HOLIDAY_MONTHS_SEARCH_RANGE * 30)
    for day in days:
        day_date = datetime.strptime(day.date, '%d/%m/%Y')
        if not (lower_bound <= day_date <= upper_bound):
            messagebox.showinfo("Warning!", f"{day.date} is outside the scope of the employee availability search! "
                                            f"The search is currently set to +/- {config.HOLIDAY_MONTHS_SEARCH_RANGE} "
                                            f"months. Consider changing HOLIDAY_MONTHS_SEARCH_RANGE in config file.")


def reset_assignments(people, days):
    for person in people:
        person.working_days = []
        person.working_saturdays -= person.assigned_saturdays
        person.assigned_saturdays = 0
        person.working_sundays -= person.assigned_sundays
        person.assigned_sundays = 0
    for day in days:
        day.assigned_people = []


def show_all_possible_assignments(days, people):
    check_dates(days)
    for day in days:
        available_people = [
            person for person in people
            if not person.exclude and day.date not in person.unavailable
        ]
        day.assigned_people = [person.initials for person in available_people]
    summary = "\n".join([f"{day.date} - {', '.join(day.assigned_people)}" for day in days])
    messagebox.showinfo("Assignments", summary)
    reset_assignments(people, days)

