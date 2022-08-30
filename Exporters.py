from abc import ABC, abstractmethod
from random import randint

import datetime
import xlsxwriter
import ics
import os

class Colors:
    def __init__(self, nb_color):
        self.colors = self.random_colors(nb_color)
    
    def random_colors(self, n):
        color = []
        for i in range(n):
            color.append('#%06X' % randint(0, 0xFFFFFF))
        return color
    
    def set_colors(self, n):
        self.colors = random_colors(n)

class Cursor:
    def __init__(self, _row, _col, _color):
        self.row = _row
        self.col = _col
        self.color = _color

    def increment_color(self, nb):
        self.color += nb

    def increment_row(self, nb):
        self.row += nb
    
    def increment_col(self, nb):
        self.col += nb

class Exporter(ABC):
    @abstractmethod
    def add_event(self, _begin, _end, _current_year, content, module_name=""):
        """Add an event to the final calendar

        :param _begin: Start of the event
        :type _begin: datetime.datetime
        :param _end: End of the event
        :type _end: datetime.datetime
        :param _current_year: Current year
        :type _current_year: str
        :param content: Content for the event (eg: title)
        :type content: str
        """
        ...

    @abstractmethod
    def export(self):
        ...


    def init_format(self, nb, *, init_data):
        ...

    def add_event_group(self, code, title):
        ...


class Calendar(Exporter):
    def __init__(self, _title):
        self.title = _title
        self.calendar = ics.Calendar()
    
    def add_event(self, _begin, _end, _current_year, _content, module_name=""):
        new_p = ics.Event()
        new_p.name = f'{module_name} - {_content}'
        new_p.summary = _content
        new_p.description = _content
        new_p.begin = _begin.strftime('%Y-%m-%d')
        new_p.end = _end.strftime('%Y-%m-%d')
        new_p.make_all_day()
        self.calendar.events.add(new_p)

    def add_event_group(self, code, title):
        pass

    def init_format(self, nb, *, init_data):
        pass
    
    def export(self):
        f = open(f'{self.title}.ics', 'w')
        f.write(self.calendar.serialize())
        f.close()


class Excel(Exporter):
    def __init__(self, title):
        self.workbook = xlsxwriter.Workbook(f'{title}.xlsx')
        self.worksheet = self.workbook.add_worksheet()
        self.worksheet.set_column(0, 0, 40)
        self.worksheet.set_column(1, 1, 15)
        self.cell_format = self.workbook.add_format({'bold': True})
        self.merge_format = self.workbook.add_format(
            {'align': 'center', 'bold': True})
        self.cursor = Cursor(2, 0, -1)
        self.colors = []
        self.worksheet.freeze_panes(0, 1)

    def write_cell_merge(self, row, col, index, data):
        tmp = data[0][index]
        switch_index = col
        for elem in data:
            if tmp != elem[index]:
                self.worksheet.merge_range(row, switch_index, row, 
                                           col - 1, tmp, self.merge_format)
                switch_index = col
                tmp = elem[index]
            self.worksheet.write(row, col, elem[index], self.cell_format)
            col += 1
        if switch_index != col-1:
            self.worksheet.merge_range(row, switch_index, row, 
                                       col - 1, tmp, self.merge_format)

    def write_cells(self, col, row, index, data):
        switch_index = 0
        for elem in data:
            self.worksheet.write(row, col, elem[index], self.cell_format)
            col += 1

    def init_header(self, dates):
        self.write_cell_merge(row=0, col=2, index=0, data=dates)
        self.write_cell_merge(row=1, col=2, index=1, data=dates)
        self.write_cells(row=2, col=2, index=2, data=dates)

    #TODO: Add to interface
    def init_format(self, nb, *, init_data):
        self.colors = Colors(nb)
        self.init_header(init_data)
        self.worksheet.freeze_panes(0, 1)


    def overlap_project(self, begin, end):
        for i in range(begin, end):
            try:
                if self.worksheet.table[self.cursor.row][i]:
                    self.cursor.increment_row(1)
                    return
            except:
                continue


    def add_event(self, begin, end, current_year, content, module_name=""):
        if self.colors == []:
            print(f'Colors are not set call the init_format function')
            #TODO: proper error handling
            return
        #TODO: read the generated spreadsheet to avoid overlaps
        color_format = self.workbook.add_format(
                {
                    'bg_color': self.colors.colors[self.cursor.color], 
                    'align': 'center'
                }
            )
        b_weeknum = int(begin.strftime('%V'))
        e_weeknum = int(end.strftime('%V'))
       
        b_offset = 0 if b_weeknum == 52 else 1
        e_offset = 0 if e_weeknum == 52 else 1
       
        b = (b_weeknum + 22) % (52 * (begin.year 
            - int(current_year) + b_offset)) + 2
        e = (e_weeknum + 22) % (52 * (end.year 
            - int(current_year) + e_offset)) + 2
       
        self.overlap_project(b, e)
        
        if b_weeknum - e_weeknum != 0:
            self.worksheet.merge_range(
                self.cursor.row, b, self.cursor.row,
                e, content, color_format)
            return
        self.worksheet.write(self.cursor.row, b, content, color_format)
        self.worksheet.set_column(b, e, len(content))

    def add_event_group(self, code, title):
        self.cursor.increment_color(1)
        self.cursor.increment_row(1)
        self.worksheet.write(self.cursor.row, self.cursor.col, title)
        self.worksheet.write(self.cursor.row, self.cursor.col + 1, code)
  
    def export(self):
        self.workbook.close()
