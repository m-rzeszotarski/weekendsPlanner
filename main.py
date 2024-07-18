from tkinter import (filedialog, messagebox, Label, Spinbox, Button, END, Toplevel, BooleanVar, Checkbutton, Tk, Frame,
                     Entry)
from tkcalendar import Calendar
from datetime import datetime, timedelta
from models import Day, Person
from sat_sun_ratio import count_weekend_days
from unavailable_check import check_people_availability
import config
import pandas
import random


def add_selected_day():
    selected_date = cal.get_date()
    # Format change to day/month/year
    date_obj = datetime.strptime(selected_date, '%Y-%m-%d')
    formatted_date = date_obj.strftime('%d/%m/%Y')
    new_day = Day(formatted_date)
    insert_sorted_day(new_day)
    update_selected_days_list()


def insert_sorted_day(new_day):
    global selected_days
    selected_days.append(new_day)
    selected_days.sort(key=lambda x: datetime.strptime(x.date, '%d/%m/%Y'))


def update_selected_days_list():
    for widget in listbox_frame.winfo_children():
        widget.destroy()

    people_entries.clear()

    for idx, day in enumerate(selected_days):
        day_label = Label(listbox_frame, text=str(day))
        day_label.grid(row=idx, column=0, padx=5, pady=2)

        people_spinbox = Spinbox(listbox_frame, from_=0, to=100, width=5)
        people_spinbox.delete(0, "end")
        people_spinbox.insert(0, str(day.number_of_people))  # Value from day object as a default
        people_spinbox.grid(row=idx, column=1, padx=5, pady=2)

        people_spinbox.config(command=lambda d=day, s=people_spinbox: update_day_people(d, s))

        people_entries.append((day, people_spinbox))

        remove_button = Button(listbox_frame, text="Remove", command=lambda d=day: remove_day(d))
        remove_button.grid(row=idx, column=2, padx=5, pady=2)


def update_day_people(day, spinbox):
    try:
        day.number_of_people = int(spinbox.get())
    except ValueError:
        day.number_of_people = 0


def remove_day(day):
    selected_days.remove(day)
    update_selected_days_list()


def open_file_explorer(entry_field):
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if file_path:
        entry_field.delete(0, END)
        entry_field.insert(0, file_path)


def process_file():
    holiday_file_path = holiday_path_entry.get()
    shifts_file_path = shifts_path_entry.get()
    if holiday_file_path and shifts_file_path:
        shifts_dict = count_weekend_days(shifts_file_path, shift_file_names)
        holiday_dict = check_people_availability(holiday_file_path)
        update_unavailable_dates(holiday_dict, people_in_team)
        update_working_weekend_days_ratio(shifts_dict, people_in_team)
        messagebox.showinfo("Processing holiday file", f"Found {len(holiday_dict)} unavailable days")
        shifts_result_text = "\n".join(
            [f"{person.shifts_name}: WSat: {person.working_saturdays}, WSun: {person.working_sundays}" for person in
             people_in_team])
        messagebox.showinfo("Processing shifts file",
                            f"{shifts_result_text}")
    else:
        messagebox.showwarning("Error", "Chose file paths")


def update_unavailable_dates(unavailable_dict, people):
    for date, unavailable_people in unavailable_dict.items():
        for person in people:
            if person.holiday_name in unavailable_people:
                person.unavailable.append(date)


def update_working_weekend_days_ratio(shifts_dict, people):
    for name, ratio_list in shifts_dict.items():
        for person in people:
            if person.shifts_name == name:
                person.working_saturdays = ratio_list[0]
                person.working_sundays = ratio_list[1]


def check_dates():
    now = datetime.now()
    lower_bound = now - timedelta(days=config.HOLIDAY_MONTHS_SEARCH_RANGE * 30)
    upper_bound = now + timedelta(days=config.HOLIDAY_MONTHS_SEARCH_RANGE * 30)
    for day in selected_days:
        day_date = datetime.strptime(day.date, '%d/%m/%Y')
        if not (lower_bound <= day_date <= upper_bound):
            messagebox.showinfo("Warning!", f"{day.date} is outside the scope of the employee availability search! "
                                            f"The search is currently set to +/- {config.HOLIDAY_MONTHS_SEARCH_RANGE} "
                                            f"months. Consider changing HOLIDAY_MONTHS_SEARCH_RANGE in config file.")


def reset_assignments():
    for person in people_in_team:
        person.working_days = []
        person.working_saturdays -= person.assigned_saturdays
        person.assigned_saturdays = 0
        person.working_sundays -= person.assigned_sundays
        person.assigned_sundays = 0
    for day in selected_days:
        day.assigned_people = []


def assign_people_to_days(days, people):
    check_dates()
    for day in days:
        available_people = [
            person for person in people
            if not person.exclude and day.date not in person.unavailable
        ]

        available_people_with_correct_ratio = [
            person for person in available_people
            if (day.day_of_week == 'Saturday' and person.working_saturdays < person.working_sundays) or
               (day.day_of_week == 'Sunday' and person.working_sundays < person.working_saturdays)
        ]

        available_people_with_correct_ratio_nw = [
            person for person in available_people_with_correct_ratio
            if not person.working_days
        ]

        available_people_nw = [
            person for person in available_people
            if not person.working_days
        ]

        assigned_people = []

        def assign_people(people_list, num_needed):
            nonlocal assigned_people
            random.shuffle(people_list)
            for person in people_list:
                if num_needed == 0:
                    break
                if not any(
                        (datetime.strptime(day.date, '%d/%m/%Y') - timedelta(days=1)).strftime(
                            '%d/%m/%Y') in working_day or
                        (datetime.strptime(day.date, '%d/%m/%Y') + timedelta(days=1)).strftime(
                            '%d/%m/%Y') in working_day
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

        assign_people(available_people_with_correct_ratio_nw, day.number_of_people)
        if len(assigned_people) < day.number_of_people:
            assign_people(available_people_nw, day.number_of_people - len(assigned_people))
        if len(assigned_people) < day.number_of_people:
            assign_people(available_people_with_correct_ratio, day.number_of_people - len(assigned_people))
        if len(assigned_people) < day.number_of_people:
            assign_people(available_people, day.number_of_people - len(assigned_people))
        if len(assigned_people) < day.number_of_people:
            while len(assigned_people) < day.number_of_people:
                random_person = random.choice(people_in_team)
                assigned_people.append(random_person)
                random_person.working_days.append(day.date)
                if day.day_of_week == 'Saturday':
                    random_person.working_saturdays += 1
                    random_person.assigned_saturdays += 1
                elif day.day_of_week == 'Sunday':
                    random_person.working_sundays += 1
                    random_person.assigned_sundays += 1
                messagebox.showwarning(
                    "Warning",
                    f"No available people at {day.date}, assigned random person!."
                )

        day.assigned_people = [person.initials for person in assigned_people]

    summary = "\n".join([f"{day.date} - {', '.join(day.assigned_people)}" for day in selected_days])
    messagebox.showinfo("Assignments", summary)
    reset_assignments()


def show_all_possible_assignments(days, people):
    check_dates()
    for day in days:
        available_people = [
            person for person in people
            if not person.exclude and day.date not in person.unavailable
        ]
        day.assigned_people = [person.initials for person in available_people]
    summary = "\n".join([f"{day.date} - {', '.join(day.assigned_people)}" for day in selected_days])
    messagebox.showinfo("Assignments", summary)
    reset_assignments()


def open_exclude_people_window():
    exclude_window = Toplevel(root)
    exclude_window.title("Exclude People")
    exclude_window.geometry("300x400")

    checkboxes = []

    for person in people_in_team:
        var = BooleanVar(value=person.exclude)
        cb = Checkbutton(exclude_window, text=person.initials, variable=var)
        cb.pack(anchor='w')
        checkboxes.append((person, var))

    def apply_changes():
        for x_person, x_var in checkboxes:
            x_person.exclude = x_var.get()
        exclude_window.destroy()

    apply_button = Button(exclude_window, text="Apply", command=apply_changes)
    apply_button.pack(pady=10)


root = Tk()
root.title("Weekend Work Planner")
root.geometry("600x800")

path_frame = Frame(root)
path_frame.pack(pady=10)

holiday_path_label = Label(path_frame, text="Holiday file:")
holiday_path_label.grid(row=0, column=0, padx=5, pady=5)
holiday_path_entry = Entry(path_frame, width=40)
holiday_path_entry.grid(row=0, column=1, padx=5, pady=5)
open_holiday_button = Button(path_frame, text="Choose file", command=lambda: open_file_explorer(holiday_path_entry))
open_holiday_button.grid(row=0, column=2, padx=5, pady=5)

shifts_path_label = Label(path_frame, text="Shifts file:")
shifts_path_label.grid(row=1, column=0, padx=5, pady=5)
shifts_path_entry = Entry(path_frame, width=40)
shifts_path_entry.grid(row=1, column=1, padx=5, pady=5)
open_shifts_button = Button(path_frame, text="Choose file", command=lambda: open_file_explorer(shifts_path_entry))
open_shifts_button.grid(row=1, column=2, padx=5, pady=5)

load_files_button = Button(path_frame, text="Load Files", command=process_file)
load_files_button.grid(row=1, column=3, padx=5, pady=5)

cal = Calendar(root, selectmode='day',
               year=datetime.now().year, month=datetime.now().month + 1,
               date_pattern='yyyy-mm-dd')  # Change of date format in calendar
cal.pack(pady=20)

selected_days = []
people_entries = []

calendar_buttons_frame = Frame(root)
calendar_buttons_frame.pack(pady=10)

exclude_button = Button(calendar_buttons_frame, text="Exclude people", command=open_exclude_people_window)
exclude_button.grid(row=0, column=0, padx=5)

add_day_button = Button(calendar_buttons_frame, text="Add selected day", command=add_selected_day)
add_day_button.grid(row=0, column=1, padx=5)


listbox_frame = Frame(root)
listbox_frame.pack(pady=20)

assign_button_frame = Frame(root)
assign_button_frame.pack(pady=10)

all_possible_assignments_button = Button(assign_button_frame, text="Show all availabilities for selected days",
                                         command=lambda: show_all_possible_assignments(selected_days, people_in_team))
all_possible_assignments_button.grid(row=0, column=0, padx=5)

assign_button = Button(assign_button_frame, text="Assign people to days",
                       command=lambda: assign_people_to_days(selected_days, people_in_team))
assign_button.grid(row=0, column=1, padx=5)

version_frame = Frame(root)
version_frame.pack(side='bottom', anchor='se', padx=10, pady=10)

version_label = Label(version_frame, text="Created by MIRZ (v1.0)")
version_label.pack(anchor='se')

try:
    df = pandas.read_csv("names_initials.csv")
    names_list = df.values.tolist()
    people_in_team = []
    for n in names_list:
        new_person = Person(n[0], n[1], n[2])
        for initials in config.DEFAULT_EXCLUDED_PEOPLE:
            if initials == new_person.initials:
                new_person.exclude = True
        people_in_team.append(new_person)
    shift_file_names = [person.shifts_name for person in people_in_team]

except FileNotFoundError:
    messagebox.showerror("Error", "File names_initials.csv was not found!")
    root.destroy()

if __name__ == '__main__':
    root.mainloop()
