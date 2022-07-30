import requests
import json
import sys
import datetime
import pandas as pd
import xlsxwriter
from random import randint
import re

# Liste tous les module avec des activités projets
# Tout les modules ayant des projets sont mis dans le tableau "Planning" avec leur projets corespondants et la date.

# vue par pédago par promo + nb credits,
# vue par mois / par pédago combien de module et combien de credits sont "créés" 

# TODO: ajouter une formule pour compter les modules de chaque responsable & les modules orphelins

# TODO: Decouper par promo avec tous les modules associés pour le variable
# Calcul pour chaque module une moyenne de réussite par pédago

token = sys.argv[1]

# dates 
currentYear = datetime.datetime.now().year - 1
pr = pd.period_range(start=f'{currentYear}-08', end=f'{currentYear+1}-11', freq='W')
dates = [(period.year, datetime.date.fromisoformat(f'{period.year}-{str(period.month).zfill(2)}-{str(period.day).zfill(2)}').strftime("%B"), period.week) for period in pr]

t = requests.get(token + f"/course/filter?format=json&preload=1&location[]=FR&location[]=FR/LYN&course[]=Code-And-Go&course[]=Dev-And-Go&course[]=bachelor/classic&course[]=premsc&course[]=webacademie&scolaryear[]={currentYear}")

result = json.loads(t.content.decode('utf-8'))
all_modules = result['items']

all_current_modules = all_modules

def write_cell_merge(worksheet, cell_format, merge_format, *, row, col, index, data):
    tmp = data[0][index]
    switch_index = col
    for elem in data:
        if tmp != elem[index]:
            worksheet.merge_range(row, switch_index, row, col - 1, tmp, merge_format)
            switch_index = col
            tmp = elem[index]
        worksheet.write(row, col, elem[index], cell_format)
        col += 1
    if switch_index != col-1:
        worksheet.merge_range(row, switch_index, row, col - 1, tmp, merge_format)

def write_cells(worksheet, cell_format, *,col, row, index, data):
    switch_index = 0
    for elem in dates:
        worksheet.write(row, col, elem[index], cell_format)
        col += 1

def write_range(worksheet, cell_format, begin, end, row, currentYear, content):
    b_weeknum = int(begin.strftime('%V'))
    e_weeknum = int(end.strftime('%V'))
    b_offset = 0 if b_weeknum == 52 else 1
    e_offset = 0 if e_weeknum == 52 else 1
    b = (b_weeknum + 22) % (52 * (begin.year - currentYear + b_offset)) + 2
    e = (e_weeknum + 22) % (52 * (end.year - currentYear + e_offset)) + 2
    if b_weeknum - e_weeknum != 0:
        worksheet.merge_range(row, b, row, e, content, cell_format)
        return
    worksheet.write(row, b, content, cell_format)

def random_colors(n):
    color = []
    for i in range(n):
        color.append('#%06X' % randint(0, 0xFFFFFF))
    return color

def clean_text(rgx_list, text):
    new_text = text
    for rgx_match in rgx_list:
        new_text = re.sub(rgx_match, '', new_text)
    return new_text

def create_sheet_headers(dates, title):
    col = 1
    row = 0
    workbook = xlsxwriter.Workbook(title+'.xlsx')
    worksheet = workbook.add_worksheet()
    cell_format = workbook.add_format({'bold': True})
    merge_format = workbook.add_format({'align': 'center', 'bold': True})
    write_cell_merge(worksheet, cell_format, merge_format, row=0, col=2, index=0, data=dates)
    write_cell_merge(worksheet, cell_format, merge_format, row=1, col=2, index=1, data=dates)
    write_cells(worksheet, cell_format, row=2, col=2, index=2, data=dates)
    worksheet.freeze_panes(0, 1)
    return workbook, worksheet

def add_project(begin, end, title, currentYear, row, cell_format, worksheet):
    begin = datetime.datetime.strptime(activity['begin'][:10], '%Y-%m-%d')
    end = datetime.datetime.strptime(activity['end'][:10], '%Y-%m-%d')
    project_text = clean_text([r"\[[\w\s]+]"], str(activity['project_title']))
    write_range(worksheet, cell_format, begin, end, row, currentYear, project_text)

def create_sheet(dates):
    row, col = 3, 0
    workbook, worksheet = create_sheet_headers(dates, 'GlobalPlanning')
    colors = random_colors(len(all_current_modules))

    for i_col, module in enumerate(all_current_modules):
        cell_format = workbook.add_format({'bg_color': colors[i_col], 'align': 'center'})
        data = requests.get(token + f"/module/{currentYear}/{module['code']}/{module['codeinstance']}/?format=json")
        activities = json.loads(data.content.decode('utf-8'))
        if len(activities['resp']) > 0:
            print(module['code'])
            print(activities['resp'][0]['login'])

        nb_acti = 0
        for activity in activities.get('activites', None):
            if activity == None:
                continue
            if activity.get('type_code') == 'proj' and activity.get('project_title') != None:
                nb_acti += 1
                add_project()

        if nb_acti > 0:
            worksheet.write(row, col, module['title'])
            worksheet.write(row, col + 1, module['code'])
            row += 1

    worksheet.set_column(0, 0, 40)
    worksheet.set_column(1, 1, 15)
    workbook.close()
    

create_sheet(dates)
