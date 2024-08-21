# Weekends Planner

This application is designed to help plan weekend work in the department of company I work for. It operates by reading two Excel files: one containing the holiday calendar with employee availability and another with the shift schedule from the previous months.
The application first identifies which employees are available on the selected weekend days. Then analyzes each employee’s working Saturday/working Sunday ratio and the total number of weekend days worked in the selected month. Based on this information, the application assigns employees to weekend shifts in the fairest way possible, ensuring an even distribution of workload.

## Packages

tkinter - This library provides the graphical user interface (GUI) for the application, allowing users to interact with the program through a visually intuitive and user-friendly interface
tkcalendar - A calendar widget specifically designed for use with Tkinter, which allows users to easily select dates within the GUI
pandas - simplifies working with CSV files
openpyxl - A library that facilitates reading and writing Excel files in the .xlsx format, enabling the application to efficiently process and update the Excel files used for scheduling and employee availability

