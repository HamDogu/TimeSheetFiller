# # Press the green button in the gutter to run the script.
# if __name__ == '__main__':
#     print_hi('PyCharm')
#
# # See PyCharm help at https://www.jetbrains.com/help/pycharm/


import tkinter as tk
import pyautogui
import numpy as np
import pandas as pd
import datetime
import os
import holidays
import keyboard  # using module keyboard
from tkinter import ttk
from tkinter import Grid
# from datetime import datetime
from tkinter.messagebox import showinfo
from calendar import month_name, calendar, month_abbr

# def keydown(e):
#     print('down', e.char)
#     hello()


root = tk.Tk()
root.resizable(False, False)
root.title('Time Sheet Filler V0.1')

# store email address and password
num_hols = tk.IntVar()
codeOne = tk.IntVar()
codeTwo = tk.IntVar()
codeThree = tk.IntVar()

# Sign in frame
signin = ttk.Frame(root)
# signin.pack(padx=10, pady=10, fill='x', expand=True)

# email
username = os.environ.get('USERNAME')
ttk.Label(text="Timesheet Filler v0.1 - " + username).grid(row=0, column=0, columnspan=2)

# email
ttk.Label(text="Number of public holidays in month:").grid(row=1, column=0, columnspan=2)
holidays_entry = ttk.Entry(signin, textvariable=num_hols)
holidays_entry.grid(row=1, column=1)
# holidays_entry.pack(fill='x', expand=True)
# holidays_entry.focus()

# label
ttk.Label(text="Please select a month:").grid(row=2, column=0)
# label.pack(fill=tk.X, padx=10, pady=5)

# create a combobox
selected_month = tk.StringVar()
month_cb = ttk.Combobox(root, textvariable=selected_month)
month_cb.grid(row=2, column=1)

# get first 3 letters of every month name
month_cb['values'] = [month_name[m][0:3] for m in range(1, 13)]

# prevent typing a value
month_cb['state'] = 'readonly'

# place the widget
# month_cb.pack(fill=tk.X, padx=10, pady=5)

# Code labels
ttk.Label(text="1st Code %:", width=10).grid(row=4, column=0)
codeOneEntry = ttk.Entry(textvariable=codeOne, width=10)
codeOneEntry.grid(row=4, column=1)
ttk.Label(text="2nd Code %:", width=10).grid(row=5, column=0)
codeTwoEntry = ttk.Entry(textvariable=codeOne, width=10)
codeTwoEntry.grid(row=5, column=1)
ttk.Label(text="3rd Code %:", width=10).grid(row=6, column=0)
codeThreeEntry = ttk.Entry(textvariable=codeOne, width=10)
codeThreeEntry.grid(row=6, column=1)

# codeOneLabel = ttk.Label(signin, text="1st Code %:")
# codeOneLabel.pack(side="left")
#
# codeOneEntry = ttk.Entry(signin, textvariable=num_hols)
# codeOneEntry.pack(side="left", fill='x', expand=True)
#
# codeTwoLabel = ttk.Label(signin, text="2nd Code %:")
# codeTwoLabel.pack(fill='x')# side="left")
#
# codeTwoEntry = ttk.Entry(signin, textvariable=num_hols)
# codeTwoEntry.pack(side="left", fill='x', expand=True)
#
# codeThreeLabel = ttk.Label(signin, text="3rd Code %:")
# codeThreeLabel.pack(side="left")
#
# codeThreeEntry = ttk.Entry(signin, textvariable=num_hols)
# codeThreeEntry.pack(side="left", fill='x', expand=True)


# def holidays():
#     #uk_holidays = []
#     for date in holidays.UnitedStates(years=2022).items():
#         print(date) #uk_holidays.append((str(date[0])))
#     #print(uk_holidays)

# bind the selected value changes
def month_changed(event):
    """ handle the month changed event """
    showinfo(
        title='Result',
        message=f'You selected {selected_month.get()}!'
    )

def hello():
    label1 = tk.Label(root, text='Hello World!', fg='green', font=('helvetica', 15, 'bold'))
    canvas1.create_window(100, 20, window=label1)
    # root.iconify()
    # holidays()
    switchWindow()


def switchWindow():
    # Holds down the alt key
    pyautogui.keyDown("alt")
    # Presses the tab key once
    pyautogui.press("tab")
    # Lets go of the alt key
    pyautogui.keyUp("alt")

def workDays(event):
    current_year = datetime.date.today().year
    from time import strptime
    current_mon = strptime(selected_month.get(),'%b').tm_mon
    mon = str(current_mon)
    next_mon = str(current_mon + 1)
    if current_mon < 10:
        mon = '0' + mon
        next_mon = '0' + next_mon

    # Experimental Holiday function
    # for holiday in holidays.Germany(years=[2020, 2021]).items():
    #     print(holiday)

    this_mon = str(current_year) + '-' + mon
    next_mon = str(current_year) + '-' + next_mon

    bus_days = np.busday_count(this_mon, next_mon)

    bus_days = int(bus_days) - int(holidays_entry.get())


    showinfo(
        title='Result',
        message=f'You selected {bus_days}!'
    )

# def fillLine():
#     # d

month_cb.bind('<<ComboboxSelected>>', workDays)

# set the current month
current_month = datetime.datetime.now().strftime('%b')
month_cb.set(current_month)

canvas1 = tk.Canvas(root, width=200, height=80)
# canvas1.bind("<KeyPress>", keydown)
canvas1.grid(row=7, column=0, columnspan=2)
# canvas1.pack()
canvas1.focus_set()

button1 = tk.Button(text='Fill me up', command=hello, bg='brown', fg='white')
canvas1.create_window(100, 60, window=button1)


root.mainloop()


