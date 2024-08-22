from tkinter import (filedialog, messagebox, Label, Spinbox, Button, END, Toplevel, BooleanVar, Checkbutton, Tk, Frame,
                     Entry, ttk)
from tkcalendar import Calendar
from datetime import datetime
from models import Day, Person
from shifts_file_read import count_weekend_days
from holiday_calendar_read import check_people_fields
from assign_people import show_all_possible_assignments, assign_people_to_days
import config
import configparser
import pandas
import os


# Add a day selected in the calendar
def add_selected_day():
    selected_date = cal.get_date()
    if selected_date == "":
        messagebox.showinfo("Error", "Please select a day from calendar")
        return
    # Format change to day/month/year
    date_obj = datetime.strptime(selected_date, '%Y-%m-%d')
    formatted_date = date_obj.strftime('%d/%m/%Y')
    if all(formatted_date not in day.date for day in selected_days):
        new_day = Day(formatted_date)
        insert_sorted_day(new_day)
        update_selected_days_list()
    else:
        messagebox.showinfo("Error", "This day has already been chosen")


# Sort selected days by date
def insert_sorted_day(new_day):
    global selected_days
    selected_days.append(new_day)
    selected_days.sort(key=lambda x: datetime.strptime(x.date, '%d/%m/%Y'))


# Update displayed selected days
def update_selected_days_list():
    for widget in listbox_frame.winfo_children():
        widget.destroy()

    people_entries.clear()

    for idx, day in enumerate(selected_days):
        day_label = Label(listbox_frame, text=str(day))
        day_label.grid(row=idx, column=0, padx=5, pady=2)

        # Select number of people to assign for selected date
        people_spinbox = Spinbox(listbox_frame, from_=0, to=100, width=5)
        people_spinbox.delete(0, "end")
        people_spinbox.insert(0, str(day.number_of_people))  # Value from day object as a default
        people_spinbox.grid(row=idx, column=1, padx=5, pady=2)

        people_spinbox.config(command=lambda d=day, s=people_spinbox: update_day_people(d, s))

        people_entries.append((day, people_spinbox))
        # Display forced people
        list_of_forced_people = []
        for person in people_in_team:
            for date in person.forced_working_days:
                formatted_person_date = datetime.strptime(date, '%d/%m/%Y')
                formatted_check_day = datetime.strptime(day.date, '%d/%m/%Y')
                if formatted_person_date == formatted_check_day:
                    list_of_forced_people.append(person.initials)
        if list_of_forced_people:
            forced_people_label = Label(listbox_frame, text=str(list_of_forced_people).replace("'", ""))
            forced_people_label.grid(row=idx, column=2, padx=5, pady=2)
        # End of new code
        manual_assignment_button = Button(listbox_frame, text="Manual assignment",
                                          command=lambda d=day: manual_assignment(d))
        manual_assignment_button.grid(row=idx, column=3, padx=5, pady=2)

        remove_button = Button(listbox_frame, text="Remove", command=lambda d=day: remove_day(d))
        remove_button.grid(row=idx, column=4, padx=5, pady=2)


# Force assignment of selected people to selected date
def manual_assignment(day):
    manual_assignment_window = Toplevel(root)
    manual_assignment_window.title(f"Manual assignment for {day.date}")
    manual_assignment_window.geometry(config.EXCLUDE_WINDOW_RESOLUTION)

    checkboxes = []

    for person in people_in_team:
        is_forced_assigned = any(day.date == forced_date for forced_date in person.forced_working_days)
        var = BooleanVar(value=is_forced_assigned)
        cb = Checkbutton(manual_assignment_window, text=person.initials, variable=var)
        cb.pack(anchor='w')
        checkboxes.append((person, var))

    def apply_changes():
        for x_person, x_var in checkboxes:
            if x_var.get():
                if day.date not in x_person.forced_working_days:
                    x_person.forced_working_days.append(day.date)
            else:
                x_person.forced_working_days = [forced_date for forced_date in x_person.forced_working_days
                                                if forced_date != day.date]

        update_selected_days_list()
        manual_assignment_window.destroy()

    apply_button = Button(manual_assignment_window, text="Apply", command=apply_changes)
    apply_button.pack(pady=10)


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


# Load holiday and shifts files
def process_file():
    holiday_file_path = holiday_path_entry.get()
    shifts_file_path = shifts_path_entry.get()
    if holiday_file_path and shifts_file_path:
        shift_file_names = [person.shifts_name for person in people_in_team]
        shifts_dict = count_weekend_days(shifts_file_path, shift_file_names)
        unavailable_people_dict = check_people_fields(holiday_file_path, config.UNAVAILABLE_WORDS_LIST)
        holiday_people_dict = check_people_fields(holiday_file_path, config.HOLIDAY_WORDS_LIST)
        try:
            update_unavailable_dates(unavailable_people_dict, people_in_team)
            update_holiday_dates(holiday_people_dict, people_in_team)
            update_working_weekend_days_ratio(shifts_dict, people_in_team)
            messagebox.showinfo("Processing holiday file", f"Found {len(unavailable_people_dict)} unavailable days")
            messagebox.showinfo("Processing holiday file", f"Found {len(holiday_people_dict)} holiday days")
        except AttributeError:
            messagebox.showinfo("Error", "No unavailable days found. Probably Holiday File was not loaded properly.")
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


def update_holiday_dates(holiday_dict, people):
    for date, holiday_people in holiday_dict.items():
        for person in people:
            if person.holiday_name in holiday_people:
                person.holiday.append(date)


def update_working_weekend_days_ratio(shifts_dict, people):
    for name, ratio_list in shifts_dict.items():
        for person in people:
            if person.shifts_name == name:
                person.working_saturdays = ratio_list[0]
                person.working_sundays = ratio_list[1]


def open_exclude_people_window():
    exclude_window = Toplevel(root)
    exclude_window.title("Exclude People")
    exclude_window.geometry(config.EXCLUDE_WINDOW_RESOLUTION)

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


def check_and_load_files():
    project_directory = os.getcwd()  # Get the current working directory
    holiday_file = config.HOLIDAY_FILE_NAME
    shifts_file = config.SHIFTS_FILE_NAME

    holiday_path = os.path.join(project_directory, holiday_file)
    shifts_path = os.path.join(project_directory, shifts_file)

    holiday_file_exists = os.path.isfile(holiday_path)
    shifts_file_exists = os.path.isfile(shifts_path)

    if holiday_file_exists:
        holiday_path_entry.insert(0, holiday_path)
    if shifts_file_exists:
        shifts_path_entry.insert(0, shifts_path)

    if holiday_file_exists and shifts_file_exists:
        result = messagebox.askyesno("Files Found", "Holiday and shifts files were found in main directory."
                                                    " Do you want to load them?")
        if result:
            process_file()


# Try to load names-initials file which allow to assign data from two Excel files to correct person
def read_initials_file():
    try:
        df = pandas.read_csv(config.NAMES_INITIALS_FILE_NAME)
        names_list = df.values.tolist()
        people_list = []
        for n in names_list:
            new_person = Person(n[0], n[1], n[2])
            for initials in config.DEFAULT_EXCLUDED_PEOPLE:
                if initials == new_person.initials:
                    new_person.exclude = True
            people_list.append(new_person)
        return people_list

    except FileNotFoundError:
        messagebox.showerror("Error", f"File '{config.NAMES_INITIALS_FILE_NAME}' was not found! "
                                      f"Software cannot operate without this file!")


def modify_initials_file():
    initials_window = Toplevel(root)
    initials_window.title("Modify Employees List")
    initials_window.geometry(config.INITIALS_WINDOW_RESOLUTION)

    # When window is closed, initials file is reloaded
    def update_people_in_team():
        global people_in_team
        people_in_team = read_initials_file()
    initials_window.protocol("WM_DELETE_WINDOW", lambda: [update_people_in_team(), initials_window.destroy()])
    df = pandas.read_csv(config.NAMES_INITIALS_FILE_NAME)

    def save_to_csv():
        df.to_csv(config.NAMES_INITIALS_FILE_NAME, index=False, encoding='utf-8-sig')

    def edit_row(row_id):
        selected_person = tree.item(row_id)
        values = selected_person['values']

        def save_changes():
            df.loc[int(row_id)] = [initials_entry.get(), holiday_entry.get(), shifts_entry.get()]
            save_to_csv()
            tree.item(row_id, values=(initials_entry.get(), holiday_entry.get(), shifts_entry.get()))
            edit_window.destroy()

        edit_window = Toplevel(initials_window)
        edit_window.title("Edit Employee")

        Label(edit_window, text="Initials:").grid(row=0, column=0, padx=10, pady=5)
        Label(edit_window, text="Holiday Name:").grid(row=1, column=0, padx=10, pady=5)
        Label(edit_window, text="Shifts Name:").grid(row=2, column=0, padx=10, pady=5)

        initials_entry = Entry(edit_window)
        initials_entry.grid(row=0, column=1, padx=10, pady=5)
        initials_entry.insert(0, values[0])

        holiday_entry = Entry(edit_window)
        holiday_entry.grid(row=1, column=1, padx=10, pady=5)
        holiday_entry.insert(0, values[1])

        shifts_entry = Entry(edit_window)
        shifts_entry.grid(row=2, column=1, padx=10, pady=5)
        shifts_entry.insert(0, values[2])

        save_button = Button(edit_window, text="Save", command=save_changes)
        save_button.grid(row=3, columnspan=2, pady=10)

    def remove_row(row_id):
        confirm = messagebox.askyesno("Confirm", "Are you sure you want to delete this person?")
        if confirm:
            df.drop(index=int(row_id), inplace=True)
            df.reset_index(drop=True, inplace=True)
            save_to_csv()
            tree.delete(row_id)
            update_treeview()

    def add_row():
        def add_new_entry():
            new_initials = initials_entry.get()
            new_holiday = holiday_entry.get()
            new_shifts = shifts_entry.get()
            if new_initials and new_holiday and new_shifts:
                df.loc[len(df)] = [new_initials, new_holiday, new_shifts]
                save_to_csv()
                update_treeview()
                add_window.destroy()
            else:
                messagebox.showwarning("Error", "All fields must be filled out")

        add_window = Toplevel(initials_window)
        add_window.title("Add New Employee")

        Label(add_window, text="Initials:").grid(row=0, column=0, padx=10, pady=5)
        Label(add_window, text="Holiday Name:").grid(row=1, column=0, padx=10, pady=5)
        Label(add_window, text="Shifts Name:").grid(row=2, column=0, padx=10, pady=5)

        initials_entry = Entry(add_window)
        initials_entry.grid(row=0, column=1, padx=10, pady=5)

        holiday_entry = Entry(add_window)
        holiday_entry.grid(row=1, column=1, padx=10, pady=5)

        shifts_entry = Entry(add_window)
        shifts_entry.grid(row=2, column=1, padx=10, pady=5)

        add_employee_button = Button(add_window, text="Add", command=add_new_entry)
        add_employee_button.grid(row=3, columnspan=2, pady=10)

    def update_treeview():
        tree.delete(*tree.get_children())
        tree.config(height=len(df))
        for i, row in df.iterrows():
            # Ensure iid is an integer or unique string
            iid = str(i)
            # Ensure values are strings
            values = (str(row.iloc[0]), str(row.iloc[1]), str(row.iloc[2]))

            tree.insert('', 'end', iid=iid, values=values)

    # Treeview config
    tree = ttk.Treeview(initials_window,
                        columns=("Initials", "Holiday Name", "Shifts Name"),
                        show='headings',
                        height=len(df)
                        )
    tree.heading("Initials", text="Initials")
    tree.heading("Holiday Name", text="Holiday Name")
    tree.heading("Shifts Name", text="Shifts Name")
    tree.grid(row=0, column=0, columnspan=3)

    update_treeview()

    add_button = Button(initials_window, text="Add New Employee", command=add_row)
    add_button.grid(row=1, column=0, pady=10)

    edit_button = Button(initials_window, text="Edit Selected", command=lambda: edit_row(tree.selection()[0]))
    edit_button.grid(row=1, column=1, pady=10)

    remove_button = Button(initials_window, text="Remove Selected", command=lambda: remove_row(tree.selection()[0]))
    remove_button.grid(row=1, column=2, pady=10)


def load_config():
    cfg = configparser.ConfigParser()
    cfg.read('config.ini', encoding='utf-8')
    return cfg


def open_edit_window(section, cfg):
    def save_changes():
        for xkey, xentry in entries.items():
            cfg.set(section, xkey, xentry.get())
        with open('config.ini', 'w') as configfile:
            cfg.write(configfile)
        edit_window.destroy()  # Close the edit window after saving changes

    edit_window = Toplevel(root)
    edit_window.title(f"Edit {section}")

    entries = {}

    table_frame = Frame(edit_window)
    table_frame.pack(pady=10)

    for i, (key, value) in enumerate(cfg.items(section)):
        Label(table_frame, text=key).grid(row=i, column=0, padx=10, pady=5, sticky='e')
        entry = Entry(table_frame, width=40)
        entry.grid(row=i, column=1, padx=10, pady=5)
        entry.insert(0, value)
        entries[key] = entry

    edit_button = Button(edit_window, text="Save Changes", command=save_changes)
    edit_button.pack(pady=10)


def open_options_window():
    options_window = Toplevel(root)
    options_window.title("Options")
    options_window.geometry(config.OPTIONS_WINDOW_RESOLUTION)

    modify_initials_file_button = Button(options_window, text="Modify Employees List", command=modify_initials_file)
    modify_initials_file_button.pack(pady=5)

    cfg = load_config()

    # Create buttons for each section in the config file
    for section in cfg.sections():
        Button(options_window, text=f"Edit {section}", command=lambda s=section: open_edit_window(s, cfg)).pack(
            pady=5)

    apply_button = Button(options_window, text="Close", command=options_window.destroy)
    apply_button.pack(pady=5)


root = Tk()
root.title(config.APP_NAME)
root.geometry(config.MAIN_WINDOW_RESOLUTION)

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

version_label = Label(version_frame, text="Created by MIRZ (v1.2)")
version_label.pack(anchor='se')

options_frame = Frame(root)
options_frame.pack(side='bottom', anchor='se', padx=10, pady=10)

options_button = Button(options_frame, text="Options", command=open_options_window)
options_button.pack(anchor='se')

# Try to load names-initials file which allow to assign data from two Excel files to correct person
people_in_team = read_initials_file()

# Check whether holiday and shift file is in the main directory
check_and_load_files()

if __name__ == '__main__':
    root.mainloop()
