# Name: Give me a brieck
# Program Objective: To help teachers mark attendance in the background using TagUI Library and other Python Libraries

from tkinter import *
import tkinter
import tagui as t
import os
import json
from datetime import datetime
from tabulate import tabulate
import csv
import codecs
import pandas as pd
import json

dir_path = os.path.dirname(os.path.realpath(__file__))
filename = dir_path + "/login.json"
entry_window = 15  # Minutes before Students become late


def inputEmail():
    fRead = open(f"{filename}", "r")
    data = json.load(fRead)
    email = data["email"]
    password = data["password"]
    fRead.close()
    t.type('//*[@type="email"]', f'{email}[enter]')
    t.wait(3)
    t.type('//*[@type="password"]', f'{password}[enter]')
    t.wait(1)
    t.keyboard("[enter]")


def checkEmail():
    fRead = open(f"{filename}", "r")
    data = json.load(fRead)
    if data["email"] == "" or data["password"] == "":
        email = t.ask("Enter Email:")
        password = t.ask("Enter Password:")
        data["email"] = email
        data["password"] = password
    fRead.close()
    fWrite = open(f"{filename}", "w")
    json.dump(data, fWrite)
    fWrite.close()
    inputEmail()


t.init(visual_automation=True, chrome_browser=True,
       headless_mode=False, turbo_mode=False)
t.url("https://teams.microsoft.com/_?lm=deeplink&lmsrc=NeutralHomePageWeb&cmpid=WebSignIn#/school/teams-grid/General?ctx=teamsGrid")
t.wait(1)
t.timeout(1)
if t.exist('//*[@type="email"]'):
    checkEmail()
if t.exist('//*[@id="oops-signout-link"]'):
    t.click('//*[@id="oops-signout-link"]')
    checkEmail()
if t.exist('//*[@class="use-app-link"]'):
    t.click('//*[@class="use-app-link"]')
t.wait(1)
name = t.ask('Enter Meeting Name: ')
t.click(name)
t.wait(2)
t.dom(
    "var reportArr = document.getElementsByClassName('attendance-report-content'); var last = reportArr[reportArr.length - 1]; last.setAttribute('id','Kaleb');")
t.wait(1)
t.click("//*[@id='Kaleb']")
t.wait(3)
t.keyboard("[win]s")
t.wait(1)
t.keyboard("meetingAttendanceReport")
t.wait(5)
t.keyboard("[right]")
t.wait(0.2)
t.keyboard("[down]")
t.wait(0.2)
t.keyboard("[down]")
t.wait(0.2)
t.keyboard("[enter]")
csvFilePath = t.clipboard()
t.keyboard("[esc]")
FILE_PATH = csvFilePath  # EDIT THIS FOR UR OWN USE!!!
# function to convert everything to standardized seconds


def convert_time(duration):
    temp_storage = []  # store the duration
    temp_storage = (duration.split(" "))  # split the differing durations
    seconds = 0
    # conversion to seconds
    for time in temp_storage:
        for time_coefficient in time:
            if time_coefficient == "h":
                seconds += (int(time.split("h")[0])*3600)
            if time_coefficient == "m":
                seconds += (int(time.split("m")[0])*60)
            if time_coefficient == "s":
                seconds += (int(time.split("s")[0]))
            else:
                continue

    return seconds


def meeting_duration(meeting_information):
    seconds_elapsed = 0
    for counter in range(3, 5):
        for variables in meeting_information[counter]:
            if variables[0] == "M":
                information_date.append(variables.split("\t")[1])
            else:
                information_time.append(variables)
    seconds_elapsed += abs(datetime.strptime(information_time[1], ' %I:%M:%S %p') - datetime.strptime(
        information_time[0], ' %I:%M:%S %p')).seconds
    # seconds_elapsed += datetime.strptime(information_date[1],'%b/%d/%Y') - datetime.strptime(information_date[0],'%b/%d/%Y') #undone for two different dates

    return seconds_elapsed


name_array = []
attendance_array = []
duration_array = []
seconds_array = []
join_time_array = []
information_date = []
information_time = []
attendance_TF = []
arr = []
meeting_information = []
count = -1
attendance_duration_percentage = 0.7
# change this to determine how much time you want the students to be in the call for

df = csv.reader(codecs.open(FILE_PATH, 'rU', 'utf-16'))
arr = []
meeting_information = []
for row in df:
    arr.append(row)
# Strip off all the data that are not important
remove = []
for i in arr:
    if not i:
        break
    remove.append(i)
remove.append([])


for k in remove:
    arr.remove(k)
    meeting_information.append(k)
arr.pop(0)

attendanceSheet = []
for test in arr:
    lol = "".join(test)
    attendanceSheet.append(lol.split("\t"))

for i in attendanceSheet:
    if i.count("Organiser") == 0 and i.count("Organizer") == 0:
        name_array.append(i[0])
        duration_array.append(i[3])

# logic to get students time of enter
for students in arr:
    join_time_array.append(students[1].split("\t")[0])
    for x in join_time_array:
        if abs(datetime.strptime(x, ' %I:%M:%S %p') - datetime.strptime(meeting_information[3][1], ' %I:%M:%S %p')).seconds > entry_window * 60:
            attendance_TF.append("False")
        else:
            attendance_TF.append("True")

# logic to determine attendance
for durations in duration_array:
    seconds_array.append(convert_time(durations))
for x in seconds_array:
    count += 1
    if x < (meeting_duration(meeting_information)*attendance_duration_percentage):
        attendance_array.append("Left Early")
    elif attendance_TF[count] == "False":
        attendance_array.append("Late")
    else:
        attendance_array.append("Present")
df = pd.DataFrame(
    {"Name": name_array, "Attendance": attendance_array, "Duration": duration_array})
table = tabulate({"Name": name_array, "Attendance": attendance_array, "Duration": duration_array},
                 headers="keys", tablefmt="psql")
top = tkinter.Tk()
top.title("Give me a brieck")
top.geometry(f"500x{max(len(name_array) * 150,400)}")
labl = Label(top, text=f"{table}")
labl.grid(column=0, row=0)
top.mainloop()
writer = pd.ExcelWriter(
    f'{dir_path}/AttendanceReport{name}.xlsx', engine='xlsxwriter')
df.to_excel(writer, sheet_name='Sheet1', index=False)
writer.save()
os.remove(FILE_PATH)
print("Done")
t.close()
exit()
