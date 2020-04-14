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
worksheet.right_to_left()
#####################################################################
date_format = workbook.add_format({'num_format': 'dd/mm/yyyy', 'align': 'left'})
date_string = data["date"]
date_object = datetime.strptime(date_string, "%y-%m-%d")

################################################################
# lists
headLines = ["תאריך", "יום", "תורן", "מפקד", "הערות"]
days = data["days"]
soldiers = data["soldiers"]
Commanders = data["Commanders"]
commanderInLine = data['commanderInLine']
soldierInLine = data["soldierInLine"]
maxBorderSize = data["maxBorderSize"]
max_range = data["max_range"]
weekendDays = data["weekendDays"] % len(days)  # for when the weekend days are bigger then the week
numberOfBlocks = data["numberOfBlocks"]


# this variable is used to help sort the columns (every 7 days there is an of set)


# # add AutoFit to the cells
def adjustThecells():
    path = os.path.abspath(
        excelName)  # for some reason (you can look it up on google excel.Workbooks.Open  function needs full path
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    wb = excel.Workbooks.Open(path)
    ws = wb.Worksheets("Sheet1")
    ws.Columns.AutoFit()
    wb.Save()
    excel.Application.Quit()


# getting the list location
def gettingListLocation(listA, elementA):
    for element in range(0, len(listA), 1):
        if elementA == listA[element]:
            return element


# adding a headline for every coloum
def AddingHeadLines(col=0, row=0):
    head_line_format = workbook.add_format()
    head_line_format.set_pattern(1)
    head_line_format.set_bg_color('yellow')
    head_line_format.set_bold()
    head_line_format.set_fg_color("black")
    for headLine in headLines:
        worksheet.write(col, row, headLine, head_line_format)
        # worksheet.write(col, row + len(headLines) + maxBorderSize, headLine, head_lineformat)
        row += 1


# adding a border
def settingBorder(col, row):
    color = workbook.add_format()
    color.set_pattern(1)
    color.set_bg_color('white')
    for i in range(0, maxBorderSize, 1):
        worksheet.write(col, len(headLines) + i + row, "", color)
        # worksheet.write(col, len(headLines)+1, "", color)


# adding a date to every section
def setDate(daysPassed, col, row=0):
    # worksheet.write(col + ofSet, gettingListLocation(headLines, "תאריך") + len(headLines) + maxBorderSize,
    #               date_object.date() + timedelta(days=col - 1), date_format)
    worksheet.write(col, gettingListLocation(headLines, "תאריך") + row,
                    date_object.date() + timedelta(days=daysPassed - 1),
                    date_format)  # why the hell sunday is the last day in the calendar?


# setting the day Itself
def setDay(day, col=0, row=0):
    # worksheet.write(col + ofSet, len(headLines) + maxBorderSize + gettingListLocation(headLines, "יום"), days[day])
    worksheet.write(col, gettingListLocation(headLines, "יום") + row, days[day])


# setting the soldiers in the rows
def setSoldeir(sold, col=0, row=0, ):
    # worksheet.write(col + ofSet, len(headLines) + maxBorderSize + gettingListLocation(headLines, "תורן"),
    #       soldiers[sold])
    worksheet.write(col, gettingListLocation(headLines, "תורן") + row, soldiers[sold])


# setting the commanders
def setCommender(comm, col=0, row=0):
    # comm2 = ((comm + int(max_range / len(days))) % len(comm))
    # comm = comm % 3
    # worksheet.write(col + ofSet - 3, gettingListLocation(headLines, "מפקד") + len(headLines) + maxBorderSize,
    #                Commanders[comm])
    worksheet.write(col, gettingListLocation(headLines, "מפקד") + row, Commanders[comm])


def CreateTheBlock(row=0, lined_soldier=0, commander_in_line=0, border=True, days_passed=0):
    ofset = 0
    commander_in_line = (commander_in_line + int(days_passed / len(days)) % len(Commanders))
    lined_soldier = (lined_soldier + days_passed - (int(days_passed / len(days)) * weekendDays)) % len(soldiers)
    day_to_begin = days_passed % len(days)
    if day_to_begin != 0:
        ofset += 1
    for columns in range(day_to_begin, max_range, 1):
        days_passed += 1
        day = columns % len(days)
        if day == 0:
            AddingHeadLines(col=columns + ofset, row=row)
            if border:
                settingBorder(col=columns + ofset, row=row)
            ofset += 1
            setCommender(comm=commander_in_line, col=columns + ofset, row=row)
            commander_in_line += 1
            commander_in_line = commander_in_line % len(Commanders)
        if border:
            settingBorder(col=columns + ofset, row=row)
        setDate(col=columns + ofset, row=row, daysPassed=days_passed)
        setDay(col=columns + ofset, day=day, row=row)
        if day < len(days) - weekendDays:
            sold = lined_soldier % len(soldiers)
            setSoldeir(sold=sold, col=columns + ofset, row=row)
            lined_soldier += 1
    return days_passed


def GetTheDayNow():
    global date_object
    weekdays = (date_object.weekday() + 1) % 7  # for some reason sunday is prefer to be the last day of the week.
    date_object = date_object - timedelta(days=weekdays)
    return weekdays  # because weekday starts with monday.


# daysPassed=0
daysPassed = GetTheDayNow()
for i in range(0, numberOfBlocks, 1):
    daysPassed = CreateTheBlock(row=(len(headLines) + maxBorderSize) * i, lined_soldier=soldierInLine,
                                commander_in_line=commanderInLine, days_passed=daysPassed, border=False)
workbook.close()
adjustThecells()
