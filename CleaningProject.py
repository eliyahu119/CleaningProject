from typing import Any, Union
from datetime import datetime
from datetime import timedelta
import xlsxwriter
from xlsxwriter.worksheet import Worksheet
import json

max_range=71

with open('soldier_file.json',encoding="utf-8") as jasonFolder:
    data=json.load(jasonFolder)
workbook = xlsxwriter.Workbook('עבודת נקיון.xlsx')
worksheet: Union[Worksheet, Any] = workbook.add_worksheet()
#####################################################################
date_format = workbook.add_format({'num_format': 'dd/mm/yyyy', 'align': 'left'})
date_string = data["date"]
date_object = datetime.strptime(date_string, "%d/%m/%Y")

################################################################
#lists
headLines= ["הערות","מפקד", "תורן", "יום","תאריך"]
days=["ראשון","שני","שלישי","רביעי","חמישי","שישי","שבת"]
soldiers=data["soldiers"]
Commanders=data["Commanders"]


#getting the list location
def gettingListLocation(listA ,elementA):
    for element in range(0,len(listA),1):
        if elementA==listA[element]:
            return element

#adding a headline for every coloum
def AddingHeadLines(col = 0,row=0):
    headLineformat = workbook.add_format()
    headLineformat.set_pattern(1)
    headLineformat.set_bg_color('yellow')
    headLineformat.set_bold()
    headLineformat.set_fg_color("black")
    settingBorder(col)
    for headLine in headLines:
        worksheet.write(col,row, headLine, headLineformat)
        worksheet.write(col, row+len(headLines)+2, headLine, headLineformat)
        row+=1

#adding a border
def settingBorder(col):
        color = workbook.add_format()
        color.set_pattern(1)
        color.set_bg_color('brown')
        worksheet.write(col, len(headLines),"" ,color)
        worksheet.write(col, len(headLines)+1, "", color)

#adding a date to every section
def setDate(col):
       worksheet.write(col+ofSet, gettingListLocation(headLines, "תאריך") + len(headLines) + 2, date_object.date()+ timedelta(days=col-1), date_format)
       worksheet.write(col+ofSet, gettingListLocation(headLines,"תאריך"), date_object.date()+ timedelta(days=col-1+max_range-1), date_format)

#setting the day Itself
def setDay(col,day):
    worksheet.write(col+ofSet,len(headLines)+2+gettingListLocation(headLines,"יום") , days[day])
    worksheet.write(col+ofSet, gettingListLocation(headLines, "יום"), days[day])

#setting the soldiers in the rows
def setSoldeir(col,sold,sold2):
    worksheet.write(col + ofSet,len(headLines) + 2 + gettingListLocation(headLines, "תורן") , soldiers[sold])
    worksheet.write(col + ofSet,gettingListLocation(headLines,"תורן") , soldiers[sold2])

#setting the commenders
def setCommender(comm):
    comm2=((comm+int(max_range/7))%3)
    comm=comm%3
    worksheet.write(col+ofSet-3,gettingListLocation(headLines,"מפקד")+len(headLines)+2,Commanders[comm])
    worksheet.write(col+ofSet-3,gettingListLocation(headLines,"מפקד"),Commanders[comm2])

AddingHeadLines()
ofSet=0  #this veriable is used to help sortthe columns (every 7 days there is an ofset)
commenderInLine=data['commander']
soldierInLine=data["soldeir"]
for col in range(1,max_range,1):
    setDate(col)
    settingBorder(col+ofSet)
    day = (col - 1) % 7
    setDay(col,day)
    if(day<4):
        sold=soldierInLine%len(soldiers)
        sold2=(soldierInLine+max_range-int(max_range/7*3)-1)%len(soldiers) #the sum of the days between the sold1 to sold 2 and their their ofset
        setSoldeir(col,sold,sold2)
        soldierInLine+=1
    if(day==6):
        setCommender(commenderInLine)
        commenderInLine+=1
        if((col/7)!=((max_range-1)/7)):
             ofSet+=1
             AddingHeadLines(col+ofSet)
        else:
            data['commander']=((commenderInLine+int(max_range/7))%3)
            data['soldeir']=(soldierInLine+max_range-int(max_range/7*2)-1)%len(soldiers)
            dateString=date_object.date()+ timedelta(days=col-1+50)
            data["date"]=dateString.strftime('%d/%m/%Y')
            print(data)


workbook.close()










