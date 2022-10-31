import datetime as dt
import win32com.client

class Dates:
    def __init__(self):
        self.datesAnnual = []
        self.datesSick = []
        self.datesStudy = []
        self.datesAll = []


def get_calendar(begin,end):
    outlook = win32com.client.Dispatch('Outlook.Application').GetNamespace('MAPI')
    calendar = outlook.getDefaultFolder(9).Items
    calendar.IncludeRecurrences = True
    calendar.Sort('[Start]')
    restriction = "[Start] >= '" + begin.strftime('%d/%m/%Y') + "' AND [END] <= '" + end.strftime('%d/%m/%Y') + "'"
    calendar = calendar.Restrict(restriction)
    return calendar

def get_appointments(dates, calendar, begin, subject_kw = None,exclude_subject_kw = None, body_kw = None):
    if subject_kw == None:
        appointments = [app for app in calendar]
    else:
        appointments = [app for app in calendar if subject_kw in app.subject]
    if exclude_subject_kw != None:
        appointments = [app for app in appointments if exclude_subject_kw not in app.subject]
    cal_busy = [app.busystatus for app in appointments]
    cal_subject = [app.subject for app in appointments]
    cal_start = [app.start for app in appointments]
    cal_end = [app.end for app in appointments]
    cal_body = [app.body for app in appointments]

    #### BUSY STUFF ####
    listBusy = []
    leaveAnnual = 0
    leaveSick = 0
    leaveStudy = 0

    dates.datesAnnual = []
    dates.datesSick = []
    dates.datesStudy = []

    # Accessing Calendar
    for x in range(len(cal_subject)):
        if cal_busy[x] == 3:
            startStr = str(cal_start[x])[0:10]
            endStr = str(cal_end[x])[0:10]
            startDate = dt.datetime.strptime(startStr, "%Y-%m-%d").date()
            endDate = dt.datetime.strptime(endStr, "%Y-%m-%d").date()
            days = (endDate - startDate).days

            if endDate == begin: continue
            print(str(cal_subject[x]) + ' : ' + str(cal_busy[x]) + " - Days: " + str(days) + ' - Start: ' + startStr + " - End: " + endStr)

            subject = str(cal_subject[x]).lower()
            # Adding to total
            if 'annual' in subject or 'holiday' in subject:
                for y in range(days):
                    dates.datesAnnual.append(startDate.day + y)
                    dates.datesAll.append((startDate.day + y, 'Annual'))
                leaveAnnual += int(days)

            if 'sick' in subject or 'ill' in subject:
                for y in range(days):
                    dates.datesSick.append(startDate.day + y)
                    dates.datesAll.append((startDate.day + y, 'Sick'))
                leaveSick += int(days)

            if 'study' in subject:
                for y in range(days):
                    dates.datesStudy.append(startDate.day + y)
                    dates.datesAll.append((startDate.day + y, 'Study'))
                leaveStudy += int(days)

    print("\r\nAnnual Leave: " + str(leaveAnnual) + " - Sick Leave: " + str(leaveSick) + " - Study Leave: " + str(leaveStudy))
    print("Annual Dates: ")
    print(dates.datesAnnual)
    print("Sick Dates: ")
    print(dates.datesSick)
    print("Study Dates: ")
    print(dates.datesStudy)
    print("All: ")
    print(dates.datesAll)

    return dates

# begin = dt.datetime(2022,7,31)
# end = dt.datetime(2022,8,28)


def scrapeOutlook(begin, end):
    dates = Dates()
    cal = get_calendar(begin, end)
    dates = get_appointments(dates, cal, begin)  # , subject_kw = 'weekly', exclude_subject_kw = 'Webcast'
    return dates

# print(result)

# result.to_excel('meeting hours.xlsx')