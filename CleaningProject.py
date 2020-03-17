from typing import Any, Union
from datetime import datetime
from datetime import timedelta
import xlsxwriter
from xlsxwriter.worksheet import Worksheet
import json
import win32com.client as win32
import os

excelName = "cleaningProject.xlsx"
with open('soldier_file.json', encoding="utf-8") as jasonFolder:
    data = json.load(jasonFolder)
workbook = xlsxwriter.Workbook(excelName)
worksheet: Union[Worksheet, Any] = workbook.add_worksheet()
#####################################################################
date_format = workbook.add_format({'num_format': 'dd/mm/yyyy', 'align': 'left'})
date_string = data["date"]
date_object = datetime.strptime(date_string, "%d/%m/%Y")

################################################################
# lists
headLines = ["הערות", "מפקד", "תורן", "יום", "תאריך"]
days = data["days"]
soldiers = data["soldiers"]
Commanders = data["Commanders"]
commanderInLine = data['commanderInLine']
soldierInLine = data["soldierInLine"]
maxBorderSize = data["maxBorderSize"]
max_range = data["max_range"]
weekendDays = data["weekendDays"]


# getting the list location
def gettingListLocation(listA, elementA):
    for element in range(0, len(listA), 1):
        if elementA == listA[element]:
            return element


# adding a headline for every coloum
def AddingHeadLines(col=0, row=0):
    headLineformat = workbook.add_format()
    headLineformat.set_pattern(1)
    headLineformat.set_bg_color('yellow')
    headLineformat.set_bold()
    headLineformat.set_fg_color("black")
    settingBorder(col)
    for headLine in headLines:
        worksheet.write(col, row, headLine, headLineformat)
        worksheet.write(col, row + len(headLines) + maxBorderSize, headLine, headLineformat)
        row += 1


# adding a border
def settingBorder(col):
    color = workbook.add_format()
    color.set_pattern(1)
    color.set_bg_color('white')
    for i in range(0, maxBorderSize, 1):
        worksheet.write(col, len(headLines) + i, "", color)
    # worksheet.write(col, len(headLines)+1, "", color)


# adding a date to every section
def setDate(col):
    worksheet.write(col + ofSet, gettingListLocation(headLines, "תאריך") + len(headLines) + maxBorderSize,
                    date_object.date() + timedelta(days=col - 1), date_format)
    worksheet.write(col + ofSet, gettingListLocation(headLines, "תאריך"),
                    date_object.date() + timedelta(days=col - 1 + max_range - 1), date_format)


# setting the day Itself
def setDay(col, day):
    worksheet.write(col + ofSet, len(headLines) + maxBorderSize + gettingListLocation(headLines, "יום"), days[day])
    worksheet.write(col + ofSet, gettingListLocation(headLines, "יום"), days[day])


# setting the soldiers in the rows
def setSoldeir(col, sold, sold2):
    worksheet.write(col + ofSet, len(headLines) + maxBorderSize + gettingListLocation(headLines, "תורן"),
                    soldiers[sold])
    worksheet.write(col + ofSet, gettingListLocation(headLines, "תורן"), soldiers[sold2])


# setting the commanders
def setCommender(comm, col):
    comm2 = ((comm + int(max_range / len(days))) % 3)
    comm = comm % 3
    worksheet.write(col + ofSet - 3, gettingListLocation(headLines, "מפקד") + len(headLines) + maxBorderSize,
                    Commanders[comm])
    worksheet.write(col + ofSet - 3, gettingListLocation(headLines, "מפקד"), Commanders[comm2])


AddingHeadLines()
ofSet = 0  # this variable is used to help sort the columns (every 7 days there is an ofset)
for col in range(1, max_range + 1, 1):
    setDate(col)
    settingBorder(col + ofSet)
    day = (col - 1) % len(days)  # because the columns start with 1
    setDay(col, day)
    if day < len(days) - weekendDays:
        sold = soldierInLine % len(soldiers)
        sold2 = (soldierInLine + max_range - int(max_range / len(days) * 3)) % len(
            soldiers)  # the sum of the days between the sold1 to sold 2 and their their ofset
        setSoldeir(col, sold, sold2)
        soldierInLine += 1
    if day == 6:
        setCommender(commanderInLine, col)
        commanderInLine += 1
        if (col / len(days)) != (max_range / len(days)):
            ofSet += 1
            AddingHeadLines(col + ofSet)
workbook.close()


def adjustThecells():
    path = os.path.abspath(
        excelName)  # for some reason (you can look it up on google excel.Workbooks.Open  function needs full path
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    wb = excel.Workbooks.Open(path)
    ws = wb.Worksheets("Sheet1")
    ws.Columns.AutoFit()
    wb.Save()
    excel.Application.Quit()


adjustThecells()
