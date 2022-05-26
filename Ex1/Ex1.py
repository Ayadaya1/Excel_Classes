#!/usr/bin/env python
# -*- coding: utf-8 -*-
# vim:fileencoding=utf-8

import xlrd
from xlrd.biffh import XL_CELL_EMPTY
import xlwt
from collections import defaultdict
from datetime import datetime

def contains (a, b):
    for i in a:
        if b[1]==i[1]:
            return True
    return False

def at(a,b):
    for i in range(0,len(a)):
        if a[i][1]==b[1]:
            return i

rb = xlrd.open_workbook("../Исходные_данные/Список-студентов_(фрагмент).xlsx", formatting_info =False);
print(rb.sheet_names())
sheet = rb.sheet_by_name("Лист1")
groups = []

for i in range (6,sheet.nrows):
    if(sheet.cell(i,30).ctype!=XL_CELL_EMPTY):
        group = [1]
        groupn = sheet.cell(i,30).value
        specialty = sheet.cell(i,32).value.split(" ")[0]
        profile = sheet.cell(i,31).value.split(" ")[0]
        ochn = sheet.cell(i,36).value
        year = int(datetime.today().year - sheet.cell(i,29).value)
        thread = specialty + " " + profile + " " + str(year)
        group.append(groupn)
        group.append(thread)
        group.append(specialty)
        group.append(profile)
        if(ochn=="очное"):
            group.append("очн")
        else:
            group.append("заочн")
        group.append(year)
        if not contains(groups, group):
            groups.append(group)
        else:
            groups[at(groups,group)][0]+=1


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
    ws.write(row,0,i[1])
    ws.write(row,1,i[2])
    ws.write(row,2,i[3])
    ws.write(row,3,i[4])
    ws.write(row,4,i[5])
    ws.write(row,5,i[6])
    ws.write(row,6,i[0])
    row+=1

wb.save("../Результаты/Work.xlsx")