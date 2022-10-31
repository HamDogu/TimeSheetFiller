import math
import random
import tkinter as tk
import pyautogui
import numpy as np
import datetime
import os
import outlookScraper as outlook
from tkinter import ttk, W
from tkinter.messagebox import showinfo
from calendar import month_name, calendar, month_abbr, monthrange
from govuk_bank_holidays.bank_holidays import BankHolidays
from time import strptime

bank_holidays = BankHolidays(locale='en')
Debug = False
numFridays = 0
intervalWait = 0

pyautogui.FAILSAFE = True

versionNum = 2.0

root = tk.Tk()
root.resizable(False, False)
root.title(f'Time Sheet Filler V{versionNum}')
# root.iconbitmap("eggplant.ico")

# store email address and password
# num_hols = tk.IntVar()
# num_hols.set(2)
codeOne = tk.IntVar()
codeOne.set(35)
codeTwo = tk.IntVar()
codeTwo.set(35)
codeThree = tk.IntVar()
codeThree.set(30)

checkFirefox = tk.IntVar()
randHoursSelected = tk.IntVar()

# Sign in frame
signin = ttk.Frame(root)
# signin.pack(padx=10, pady=10, fill='x', expand=True)

# Username
username = os.environ.get('USERNAME')
ttk.Label(text=f'Timesheet Filler v{versionNum} - {username}').grid(row=0, column=0, columnspan=2)

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
    # checkDate()
    global leaveTaken
    v = tk.StringVar()
    v.set("Filling")
    label1 = tk.Label(root, textvariable=v, fg='orange', font=('helvetica', 15, 'bold'))
    canvas1.create_window(100, 20, window=label1)
    # root.iconify()

    # Running functions
    workDays()
    leaveTaken = leaveDates()
    percentCheck()
    hourCalcRand()
    switchWindow()

    # Filling and updating sheet
    if fillSheet() == True:
        v.set("Filled!")
        label1.config(fg='green')
    else:
        v.set("Failsafe Called")
        label1.config(fg='red')
    switchWindow()


def switchWindow():
    # Holds down the alt key
    pyautogui.keyDown("alt")
    # Presses the tab key once
    pyautogui.press("tab")
    # Lets go of the alt key
    pyautogui.keyUp("alt")


def workDays():  # event):
    global this_mon
    global next_mon
    current_year = datetime.date.today().year
    current_mon = strptime(selected_month.get(), '%b').tm_mon
    mon = str(current_mon)
    next_mon = str(current_mon + 1)
    prev_mon = str(current_mon - 1)
    if current_mon < 9:
        next_mon = '0' + next_mon
    if current_mon < 10:
        mon = '0' + mon
    elif current_mon == 1:
        prev_mon = 12

    # Bank Holidays
    global bankHols
    bankHols = []
    for bank_holiday in bank_holidays.get_holidays('england-and-wales', datetime.date.today().year):
        if bank_holiday['date'].month == current_mon:
            bankHols.append(bank_holiday['date'].day) # gets the day values of the bank holidays
    print(bankHols)

    # Setting month variables (badly coded lol oh well)

    this_mon = str(current_year) + '-' + mon
    if current_mon == 12:
        next_mon = str(current_year + 1) + '-01'
    else:
        next_mon = str(current_year) + '-' + next_mon
    
    global bus_days
    bus_days = np.busday_count(this_mon, next_mon)

    global numFridays
    numFridays = np.busday_count(this_mon, next_mon,weekmask='Fri')

    start_day = datetime.date(current_year, current_mon, 1)
    final_day_prev = monthrange(current_year, int(prev_mon))[1]
    final_day_current = monthrange(current_year, int(current_mon))[1]
    global outlookStart
    outlookStart = datetime.date(current_year, int(prev_mon), final_day_prev)
    global outlookEnd
    outlookEnd = datetime.date(current_year, int(current_mon), final_day_current)

    global first_day
    first_day = (start_day.weekday())#"%A"))

    global date_first_weekday
    date_first_weekday = np.busday_offset(start_day, 0, roll='forward', weekmask='Mon Tue Wed Thu Fri')
    # monthrange(current_year, int(current_mon))[0] # for first weekday of month

    print(f'First day: {str(first_day)}, Date_first_weekday: {str(date_first_weekday)}, Current month: {str(current_mon)}')

    bus_days = int(bus_days) - int(len(bankHols))


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

def hoursLeaveFri():
    first_elms = [x[0] for x in dates.datesAll]
    numFri = 0
    for i in first_elms:
        if i < 10:
            datey = this_mon + '-0' + str(i)
        else:
            datey = this_mon + '-' + str(i)
        if np.is_busday(np.datetime64(datey), weekmask="Fri"):
            numFri = numFri + 1
    return numFri

def dateToDateTime(i):
    year = datetime.date.today().year
    mon = strptime(selected_month.get(), '%b').tm_mon
    datey = datetime.datetime(year, mon, i)
    return datey


def hourCalcRand():
    # Fridays and weekdays
    print("fridays: " + str(numFridays))
    hours_weekday = (bus_days - numFridays) * 7.5
    hours_fri = numFridays * 6.5
    hours_total = hours_fri + hours_weekday

    # TODO: Change to datesAll to allow for separate logging of study leave
    # Check which of these are annual leave
    hours_leave = len(dates.datesAnnual)  # .datesAll)
    print(len(dates.datesAnnual))  # .datesAll))
    hours_leave_fri = hoursLeaveFri()
    hours_leave = hours_leave - hours_leave_fri
    # hours_total = hours_total - (hours_leave * 7.5 + hours_leave_fri * 6.5)

    h1 = round(hours_total*(pOne/100))
    h2 = round(hours_total*(pTwo/100))
    h3 = round(hours_total*(pThree/100))

    global hours1, hours2, hours3, actual_days
    hours1 = []
    hours2 = []
    hours3 = []
    actual_days = []

    i = first_day
    j = date_first_weekday.astype(datetime.datetime)
    j = j.day

    if i == 5 or i == 6:
        i = 0

    for x in range(bus_days):

        # Working our hours per day taking into account bank hols and working days per month
        if j in bankHols:  # skips the bank holiday if the days match the j value
            i = i+1
        # MaxH is for maximum hours in a day, weekday is 7.5, Friday is 6.5 and Weekend is 0
        maxH = 7.5
        if i == 4:
            maxH = 6.5
        elif i == 5 or i == 6:
            i = 0
            j = j + 2

        # Re-check if bank hol is at start of week
        if j in bankHols:  # skips the bank holiday if the days match the j value
            i = i + 1
        else:
            actual_days.append(j)



        # Setting dates to Zero if in leave
        if j in dates.datesAnnual:
            hours1.append(0)
            hours2.append(0)
            hours3.append(0)

        # Random Hours Calculations
        elif randHoursSelected.get() == 1:
            rand1 = 0
            if h1 > 0:
                rand1 = random.randint(1, 5)
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

        # If not selecting random values
        else:
            hCalc1 = round((maxH * (pOne / 100)), 2)
            hCalc2 = round((maxH * (pTwo / 100)), 2)
            hCalc3 = round(maxH - hCalc1 - hCalc2, 2)
            if h1 > 0:
                hours1.append(hCalc1)
                h1 = h1 - hCalc1
            else:
                hours1.append(0)
            if h2 > 0:
                hours2.append(hCalc2)
                h2 = h2 - hCalc2
            else:
                hours2.append(0)
            if h3 > 0:
                hours3.append(hCalc3)
                h3 = h3 - hCalc3
            else:
                hours3.append(0)

        # Moving to next day
        i = i + 1
        j = j + 1  # increasing days by 1 after hours calculation
        
    actual_days.append(j)
    # for x in range(len(hours1)):
    #     print("Hours 1: " + str(hours1[x]))
    #
    # for x in range(len(hours2)):
    #     print("Hours 2: " + str(hours2[x]))
    #
    # for x in range(len(hours3)):
    #     print("Hours 3: " + str(hours3[x]))

    print("")
    print("H1: " + str(h1))
    print("H2: " + str(h2))
    print("H3: " + str(h3))


def fillSheet():
    depCode = "Department code: 09eswm00"
    # Hours
    try:
        for i in range(bus_days):
            pyautogui.typewrite(str(hours1[i]), interval=intervalWait)
            pyautogui.press("tab")

        for i in range(2):
            pyautogui.press("tab")

        # Hours 2
        for i in range(bus_days):
            pyautogui.typewrite(str(hours2[i]), interval=intervalWait)
            pyautogui.press("tab")

        finalTabs = 3
        if checkFirefox.get() == 1:
            finalTabs = 4
        for i in range(finalTabs):
            pyautogui.press("tab")

        # Hours 3
        for i in range(bus_days):
            pyautogui.typewrite(str(hours3[i]), interval=intervalWait)
            pyautogui.press("tab")

        # Adding department code at the last box
        pyautogui.hotkey('ctrl', 'a')  # copy all in last section
        pyautogui.typewrite(depCode, interval=intervalWait)

        if leaveTaken:
            # Other section
            for i in range(bus_days):
                pyautogui.press("tab")

            for i in range(finalTabs):
                pyautogui.press("tab")

            # Sickness section
            for i in range(bus_days):
                pyautogui.press("tab")

            # Annual section
            for i in range(len(actual_days)):
                if actual_days[i] in dates.datesAnnual:
                    datey = dateToDateTime(actual_days[i])
                    hour = "7.5"
                    if datey.weekday() == 4:
                        hour = "6.5"
                    pyautogui.typewrite(hour, interval=intervalWait)
                pyautogui.press("tab")

        return True

    except pyautogui.FailSafeException:  # as e syntax added in ~python2.5
        print("Fail Safe Activated")
        #v.set("FailSafe Called")
        return False


def leaveDates():
    global dates
    dates = outlook.scrapeOutlook(outlookStart, outlookEnd)
    if len(dates.datesAll) > 0:
        return True
    return False


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

    finalTabs = 3
    if checkFirefox.get() == 1:
        finalTabs = 4
    print("Final tabs:" + str(finalTabs))

    for i in range(finalTabs):
        pyautogui.press("tab")

    for i in range(bus_days):
        pyautogui.typewrite(['backspace'], interval=intervalWait)
        pyautogui.press("tab")

    # Department code wipe
    pyautogui.hotkey('ctrl', 'a')  # copy all in last section
    pyautogui.typewrite(['backspace'], interval=intervalWait)

def checkDate():
    if datetime.date.today() > datetime.date(2022, 7, 1):
        showinfo(
            title='Version Out of Date :(',
            message=f'Current version of Time Sheet Filler is no longer supported, please contact your local Hamid for an upgrade. Thank you for using our service :)'
        )
        exit()


month_cb.bind('<<ComboboxSelected>>')
# month_cb.bind('<<ComboboxSelected>>', workDays)

# set the current month
current_month = datetime.datetime.now().strftime('%b')
month_cb.set(current_month)

ttk.Label(text="Firefox:", width=11, justify= "left").grid(row=8, column=0)
tk.Checkbutton(root, variable=checkFirefox, onvalue=1, offvalue=0).grid(row=8, column=1, sticky=W)

ttk.Label(text="Rand Hrs:", width=11, justify= "left").grid(row=9, column=0)
tk.Checkbutton(root, variable=randHoursSelected, onvalue=1, offvalue=0).grid(row=9, column=1, sticky=W)

canvas1 = tk.Canvas(root, width=200, height=110)
# canvas1.bind("<KeyPress>", keydown)
canvas1.grid(row=10, column=0, columnspan=2)
# canvas1.pack()
canvas1.focus_set()

button1 = tk.Button(text='Fill me up', command=startFill, bg='blue', fg='white')
canvas1.create_window(100, 60, window=button1)

button1 = tk.Button(text='Clear Sheet', command=deleteSheet, bg='red', fg='white')
canvas1.create_window(100, 90, window=button1)


root.mainloop()