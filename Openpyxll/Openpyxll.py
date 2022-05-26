#!/usr/bin/env python
# -*- coding: utf-8 -*-
# vim:fileencoding=utf-8

from openpyxl import load_workbook, workbook
from datetime import datetime
from openpyxl import Workbook
from openpyxl.descriptors.base import DateTime
import os

def contains (a, b):
    for i in a:
        if b[1]==i[1]:
            return True
    return False

def at(a,b):
    for i in range(0,len(a)):
        if a[i][1]==b[1]:
            return i

groups = []

wb = load_workbook("../Исходные_данные/Список-студентов_(фрагмент).xlsx",read_only = True)
ws = wb['Лист1']
for row in ws.rows:
    if row[30].value and row[30].value!="Группа":
        group = [1]
        groupn = row[30].value
        specialty = row[32].value.split(" ")[0]
        profile = row[31].value.split(" ")[0]
        ochn = row[36].value
        year = int(datetime.today().year - int(row[29].value))
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
        print(groups[-1])

workbook = Workbook()
sheet = workbook.active
sheet["A1"] = "Группа"
sheet["B1"] = "Поток"
sheet["C1"] = "Специальность/\nНаправление"
sheet["D1"] = "Специализация/\nПрофиль"
sheet["E1"] = "Очн/\nЗаочн"
sheet["F1"] = "Год набора"
sheet["G1"] = "Кол-во\nстудентов"

nrow = 2
for group in groups:
    sheet.cell(row = nrow, column = 1).value = group[1]
    sheet.cell(row = nrow, column = 2).value = group[2]
    sheet.cell(row = nrow, column = 3).value = group[3]
    sheet.cell(row = nrow, column = 4).value = group[4]
    sheet.cell(row = nrow, column = 5).value = group[5]
    sheet.cell(row = nrow, column = 6).value = group[6]
    sheet.cell(row = nrow, column = 7).value = group[0]
    nrow+=1
workbook.save(filename = "../Результаты/work_openpyxl.xlsx")
wb.close()