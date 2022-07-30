import xlsxwriter
import requests
import json
import datetime
import re
from random import randint
import pandas as pd


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

def get_activities(token, current_year, module):
    data = requests.get(token + f"/module/{current_year}/{module['code']}/{module['codeinstance']}/?format=json")
    activities = json.loads(data.content.decode('utf-8'))
    return activities

def get_modules(token, current_year):
    t = requests.get(token + f"/course/filter?format=json&preload=1&location[]=FR&location[]=FR/LYN&course[]=Code-And-Go&course[]=Dev-And-Go&course[]=bachelor/classic&course[]=premsc&course[]=webacademie&scolaryear[]={current_year}")
    result = json.loads(t.content.decode('utf-8'))
    return result['items']

def clean_text(rgx_list, text):
    new_text = text
    for rgx_match in rgx_list:
        new_text = re.sub(rgx_match, '', new_text)
    return new_text

class Colors:
    def __init__(self, nb_color):
        colors = self.random_colors(nb_color)
    
    def random_colors(self, n):
        color = []
        for i in range(n):
            color.append('#%06X' % randint(0, 0xFFFFFF))
        return color
    
    def set_colors(self, n):
        self.colors = random_colors(n)

class Cursor:
    def __init__(self, r, c):
        self.row = r
        self.col = c

# TODO add intranet class composition to allow this class to call intranet abstraction
class Planning:
    def __init__(self, title):
        self.workbook = xlsxwriter.Workbook(title+'.xlsx')
        self.worksheet = self.workbook.add_worksheet()
        self.cell_format = self.workbook.add_format({'bold': True})
        self.merge_format = self.workbook.add_format({'align': 'center', 'bold': True})
        # TODO: calculate proper scholar year
        self.current_year =  datetime.datetime.now().year - 1 
        pr = pd.period_range(start=f'{self.current_year}-08', end=f'{self.current_year+1}-11', freq='W')
        self.dates = [(period.year, datetime.date.fromisoformat(f'{period.year}-{str(period.month).zfill(2)}-{str(period.day).zfill(2)}').strftime("%B"), period.week) for period in pr]
        self.all_modules = get_modules(token, self.current_year)
        self.colors = Colors(len(self.all_modules))
        self.cursor = Cursor(3, 0)
        self.worksheet.set_column(0, 0, 40)
        self.worksheet.set_column(1, 1, 15)

    def add_project(self, begin, end, title, color_format):
        """
        begin: str 
        end: str 
        title: str
        """
        _begin = datetime.datetime.strptime(begin, '%Y-%m-%d')
        _end = datetime.datetime.strptime(end, '%Y-%m-%d')
        project_text = clean_text([r"\[[\w\s]+]"], title)
        write_range(self.worksheet, self.cell_format, _begin, _end, self.cursor.row, self.current_year, project_text)

    def add_all_modules(self, modules):
        """Add a list of modules to the planning
        """
        for i_color, module in enumerate(modules):
            color_format = self.workbook.add_format({'bg_color': colors[i_color], 'align': 'center'})
            activities = get_activities(token, self.current_year, module)
            self.add_module(activites, color_format, module['title'], module['code'])
            
    def add_module(self, activites, color_format, title, code):
        """Add one module to the planning
        """
        nb_acti = 0
        for activity in activities.get('activites', None):
            if activity == None:
                continue
            if activity.get('type_code') == 'proj' and activity.get('project_title') != None:
                nb_acti += 1
                self.add_project(activity['begin'][:10], activity['end'][:10], str(activity['project_title']), color_format)
        if nb_acti > 0:
            #TODO create a writer wrapper to allow changing output to ical
            self.worksheet.write(self.cursor.row, self.cursor.col, module['title'])
            self.worksheet.write(self.cursor.row, self.cursor.col + 1, module['code'])
            self.cursor.row += 1

    def add_header(self, dates):
        write_cell_merge(self.worksheet, self.cell_format, self.merge_format, row=0, col=2, index=0, data=dates)
        write_cell_merge(self.worksheet, self.cell_format, self.merge_format, row=1, col=2, index=1, data=dates)
        write_cells(self.worksheet, self.cell_format, row=2, col=2, index=2, data=dates)
        self.worksheet.freeze_panes(0, 1)

p = Planning('test')
p.workbook.close()