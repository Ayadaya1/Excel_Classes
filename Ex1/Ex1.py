#!/usr/bin/env python
# -*- coding: utf-8 -*-
# vim:fileencoding=utf-8

import xlrd
from xlrd.biffh import XL_CELL_EMPTY
import xlwt
from collections import defaultdict
from datetime import datetime

class Group:
    group = None
    specialty = None
    profile = None
    ochn = None
    year = None
    thread = None
    amount = None
    def __init__(self,gr, Specialty, Profile, Ochn, Year, Thread, Amount):
        self.group = gr
        self.specialty = Specialty
        self.profile = Profile
        self.ochn = Ochn
        self.year = Year
        self.thread = Thread
        self.amount = Amount

def contains (a, b):
    for i in a:
        if b.group==i.group:
            return True
    return False

def at(a,b):
    for i in range(0,len(a)):
        if a[i].group==b.group:
            return i

rb = xlrd.open_workbook("../Исходные_данные/Список-студентов_(фрагмент).xlsx", formatting_info =False);
print(rb.sheet_names())
sheet = rb.sheet_by_name("Лист1")
groups = []

for i in range (6,sheet.nrows):
    if(sheet.cell(i,30).ctype!=XL_CELL_EMPTY):
        amount = 1
        groupn = sheet.cell(i,30).value
        specialty = sheet.cell(i,32).value.split(" ")[0]
        profile = sheet.cell(i,31).value.split(" ")[0]
        ochn = sheet.cell(i,36).value
        year = int(datetime.today().year - sheet.cell(i,29).value)
        thread = specialty + " " + profile + " " + str(year)
        
        if(ochn=="очное"):
            ochn = "очн"
        else:
            ochn = "заочн"
        group = Group(groupn, specialty,profile,ochn,year,thread,amount)
        if not contains(groups, group):
            groups.append(group)
        else:
            groups[at(groups,group)].amount+=1

wb = xlwt.Workbook()
ws = wb.add_sheet("Группы")
style0 = xlwt.easyxf('font: bold true', num_format_str="#0")
ws.write(0,0, "Группа",style0)
ws.write(0,1, "Поток",style0)
ws.write(0,2, "Специальность/\nНаправление",style0)
ws.write(0,3, "Специализация/\nПрофиль",style0)
ws.write(0,4, "Очн/\nЗаочн",style0)
ws.write(0,5, "Год набора",style0)
ws.write(0,6, "Кол-во\nстудентов",style0)

row = 1

for i in groups:
    ws.write(row,0,i.group)
    ws.write(row,1,i.thread)
    ws.write(row,2,i.specialty)
    ws.write(row,3,i.profile)
    ws.write(row,4,i.ochn)
    ws.write(row,5,i.year)
    ws.write(row,6,i.amount)
    row+=1

wb.save("../Результаты/Work.xlsx")