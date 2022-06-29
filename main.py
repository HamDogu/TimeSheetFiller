# # Press the green button in the gutter to run the script.
# if __name__ == '__main__':
#     print_hi('PyCharm')
#
# # See PyCharm help at https://www.jetbrains.com/help/pycharm/
import math
import random
import tkinter as tk
import pyautogui
import numpy as np
# import pandas as pd
import datetime
import os
# import holidays
# import keyboard  # using module keyboard
from tkinter import ttk
# from tkinter import Grid
# from datetime import datetime
from tkinter.messagebox import showinfo
from calendar import month_name, calendar, month_abbr

# def keydown(e):
#     print('down', e.char)
#     hello()

jubileeMode = True
Debug = False
numFridays = 0
intervalWait = 0

pyautogui.FAILSAFE = True

root = tk.Tk()
root.resizable(False, False)
root.title('Time Sheet Filler V0.2')

# store email address and password
num_hols = tk.IntVar()
num_hols.set(2)
codeOne = tk.IntVar()
codeOne.set(33)
codeTwo = tk.IntVar()
codeTwo.set(33)
codeThree = tk.IntVar()
codeThree.set(34)

# Sign in frame
signin = ttk.Frame(root)
# signin.pack(padx=10, pady=10, fill='x', expand=True)

# email
username = os.environ.get('USERNAME')
ttk.Label(text="Timesheet Filler v0.2 - " + username).grid(row=0, column=0, columnspan=2)

# email
ttk.Label(text="Number of public holidays in month:").grid(row=1, column=0, columnspan=2)
holidays_entry = ttk.Entry(textvariable=num_hols, justify="center")
holidays_entry.grid(row=2, column=0, columnspan=2)
# holidays_entry.pack(fill='x', expand=True)
# holidays_entry.focus()

# label
ttk.Label(text="Please select a month:").grid(row=3, column=0, columnspan=2)
# label.pack(fill=tk.X, padx=10, pady=5)

# create a combobox
selected_month = tk.StringVar()
month_cb = ttk.Combobox(root, textvariable=selected_month, justify="center")
month_cb.grid(row=4, column=0, columnspan=2)

# get first 3 letters of every month name
month_cb['values'] = [month_name[m][0:3] for m in range(1, 13)]

# prevent typing a value
month_cb['state'] = 'readonly'

# place the widget
# month_cb.pack(fill=tk.X, padx=10, pady=5)

# Code labels
ttk.Label(text="1st Code %:", width=11, justify= "left").grid(row=5, column=0)
codeOneEntry = ttk.Entry(textvariable=codeOne, width=10)
codeOneEntry.grid(row=5, column=1)

ttk.Label(text="2nd Code %:", width=11, justify= "left").grid(row=6, column=0)
codeTwoEntry = ttk.Entry(textvariable=codeTwo, width=10)
codeTwoEntry.grid(row=6, column=1)

ttk.Label(text="3rd Code %:", width=11, justify= "left").grid(row=7, column=0)
codeThreeEntry = ttk.Entry(textvariable=codeThree, width=10)
codeThreeEntry.grid(row=7, column=1)


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


def startFill():
    checkDate()
    v = tk.StringVar()
    v.set("Filling")
    label1 = tk.Label(root, textvariable=v, fg='green', font=('helvetica', 15, 'bold'))
    canvas1.create_window(100, 20, window=label1)
    # root.iconify()
    # holidays()
    workDays()
    switchWindow()
    fillSheet()
    v.set("Filled!")

def switchWindow():
    # Holds down the alt key
    pyautogui.keyDown("alt")
    # Presses the tab key once
    pyautogui.press("tab")
    # Lets go of the alt key
    pyautogui.keyUp("alt")


def workDays():  # event):
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

    global bus_days
    bus_days = np.busday_count(this_mon, next_mon)

    global numFridays
    numFridays = np.busday_count(this_mon, next_mon,weekmask='Fri')

    start_day = datetime.date(current_year, current_mon, 1)
    global first_day
    first_day = (start_day.weekday())#"%A"))

    bus_days = int(bus_days) - int(holidays_entry.get())

    percentCheck()
    hourCalc()
    # showinfo(
    #     title='Result',
    #     message=f'You selected {bus_days}!'
    # )


def percentCheck():
    global  pOne, pTwo, pThree
    pOne = int(codeOneEntry.get())
    pTwo = int(codeTwoEntry.get())
    pThree = int(codeThreeEntry.get())

    if (100 - pOne - pTwo - pThree) != 0 and not Debug:
        showinfo(
            title='% Error',
            message=f'Percentages do not equal 100, do you mean 3rd% to be {(100 - pOne - pTwo)}?'
        )


def hourCalc():
    print("fridays: " + str(numFridays))
    hours_weekday = (bus_days - numFridays) * 7.5
    hours_fri = numFridays * 6.5
    hours_total = hours_fri + hours_weekday
    h1 = round(hours_total*(pOne/100))
    h2 = round(hours_total*(pTwo/100))
    h3 = round(hours_total*(pThree/100))

    global hours1, hours2, hours3
    hours1 = []
    hours2 = []
    hours3 = []

    i = first_day
    j = 0
    for x in range(bus_days):
        j = j + 1
        if jubileeMode:
            if j == 2: #
                i = 0
        # MaxH is for maximum hours in a day, weekday is 7.5, Friday is 6.5 and Weekend is 0
        maxH = 7.5
        if i == 4:
            maxH = 6.5
        elif i == 5 or i == 6:
            i = 0
        i = i + 1

        rand1 = 0
        if h1 > 0:
            rand1 = random.randint(3, 5)
            if h1 < rand1:
                hours1.append(h1)
                rand1 = round(h1)
                h1 = 0
            else:
                hours1.append(rand1)
                h1 = h1 - rand1
        else:
            hours1.append(0)

        # hours2.append(rand2)
        # h2 = h2 - rand2
        rand2 = 0
        if h2 > 0:
            rand2 = random.randint(1, (math.floor(maxH) - rand1))
            if h2 < rand2:
                hours2.append(h2)
                rand2 = round(h2)
                h2 = 0
            else:
                hours2.append(rand2)
                h2 = h2 - rand2
        else:
            hours2.append(0)

        hours3.append(maxH - rand1 - rand2)
        h3 = h3 - (maxH - rand1 - rand2)

    for x in hours1:
        print("Hours 1: " + str(hours1[x]))

    for x in hours2:
        print("Hours 2: " + str(hours2[x]))

    for x in range(len(hours3)):
        print("Hours 3: " + str(hours3[x]))

    print("")
    print("H1: " + str(h1))
    print("H2: " + str(h2))
    print("H3: " + str(h3))


def fillSheet():
    # Hours
    for i in range(bus_days):
        pyautogui.typewrite(str(hours1[i]), interval=intervalWait)
        pyautogui.press("tab")

    for i in range(2):
        pyautogui.press("tab")

    # Hours 2
    for i in range(bus_days):
        pyautogui.typewrite(str(hours2[i]), interval=intervalWait)
        pyautogui.press("tab")

    for i in range(4):
        pyautogui.press("tab")

    # Hours 3
    for i in range(bus_days):
        pyautogui.typewrite(str(hours3[i]), interval=intervalWait)
        pyautogui.press("tab")


def deleteSheet():
    workDays()
    switchWindow()

    for i in range(bus_days):
        pyautogui.typewrite(['backspace'], interval=intervalWait)
        pyautogui.press("tab")

    for i in range(2):
        pyautogui.press("tab")

    for i in range(bus_days):
        pyautogui.typewrite(['backspace'], interval=intervalWait)
        pyautogui.press("tab")

    for i in range(4):
        pyautogui.press("tab")

    for i in range(bus_days):
        pyautogui.typewrite(['backspace'], interval=intervalWait)
        pyautogui.press("tab")


def checkDate():
    if datetime.date.today() > datetime.date(2022, 7, 1):
        showinfo(
            title='Version Out of Date :(',
            message=f'Current version of Time Sheet Filler is no longer supported, please contact your local Hamid for an upgrade. Thank you for using our service :)'
        )
        exit()

# def fillLine():
#     # d


month_cb.bind('<<ComboboxSelected>>')
# month_cb.bind('<<ComboboxSelected>>', workDays)

# set the current month
current_month = datetime.datetime.now().strftime('%b')
month_cb.set(current_month)

canvas1 = tk.Canvas(root, width=200, height=110)
# canvas1.bind("<KeyPress>", keydown)
canvas1.grid(row=8, column=0, columnspan=2)
# canvas1.pack()
canvas1.focus_set()

button1 = tk.Button(text='Fill me up', command=startFill, bg='blue', fg='white')
canvas1.create_window(100, 60, window=button1)

button1 = tk.Button(text='Clear Sheet', command=deleteSheet, bg='red', fg='white')
canvas1.create_window(100, 90, window=button1)


root.mainloop()


