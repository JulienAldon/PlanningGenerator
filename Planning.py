import xlsxwriter
import requests
import json
import datetime
import re
import pandas as pd

def clean_text(rgx_list, text):
    new_text = text
    for rgx_match in rgx_list:
        new_text = re.sub(rgx_match, '', new_text)
    return new_text

class Planning:
    def __init__(self, _exporter, _intranet, month_range=('08', '11'), pedago=[]):
        self.exporter = _exporter
        self.intranet = _intranet
        # self.current_year = self.intranet.get_current_scholar_year() 
        self.current_year = 2021
        pr = pd.period_range(start=f'{self.current_year}-{month_range[0]}', end=f'{int(self.current_year)+1}-{month_range[1]}', freq='W')
        dates = [(period.year, datetime.date.fromisoformat(f'{period.year}-{str(period.month).zfill(2)}-{str(period.day).zfill(2)}').strftime("%B"), period.week) for period in pr]
        self.all_modules = self.get_modules_by_pedago(self.intranet.get_modules(self.current_year), pedago)
        self.exporter.init_format(len(self.all_modules), init_data=dates)
        self.add_all_modules(self.all_modules)
        self.exporter.export()

    def get_modules_by_pedago(self, modules, referee):
        """Filter modules by referee

        :param modules: All the modules to filter
        :type modules: List[dict]
        :param referee: All the referee to filter
        :type referee: List[str]
        """
        tmp_modules = []
        for module in modules:
            a = self.intranet.get_module(self.current_year, 
                                         module['code'], 
                                         module['codeinstance'])
            resp = [resp for resp in a['resp'] if resp['login'] in referee]
            if len(resp) > 0:
                tmp_modules.append(module)
        return tmp_modules

    def add_project(self, begin, end, title, module_name):
        """Add a project to the planning to the current row
        
        :param begin: Date of the start of the project (format %Y-%m-%d)
        :type begin: str
        :param end: Date of the end of the project (format %Y-%m-%d)
        :type end: str 
        :param title: Title of the project
        :type title: str
        :param module_name: Name of the module
        :type module_name: str

        """
        _begin = datetime.datetime.strptime(begin, '%Y-%m-%d')
        _end = datetime.datetime.strptime(end, '%Y-%m-%d')
        # project_text = clean_text([r"\[[\w\s]+]"], title)
        self.exporter.add_event(_begin, _end,
                                self.current_year, title, module_name)

    def add_all_modules(self, modules):
        """Add a list of modules to the planning
        
        :param modules: All the modules to add
        :type modules: List
        """
        for module in modules:
            if module['semester'] == 0: #Skip semester 0 modules
                continue
            activities = self.intranet.get_activities(self.current_year,
                                                      module['code'],
                                                      module['codeinstance'])
            self.add_module(activities, module['title'], module['code'])
            
    def add_module(self, activities, title, code):
        """Add one module to the planning

        :param activites: All activities of the module
        :type activites: List
        :param title: Title of the module
        :type title: str
        :param code: Code of the module
        :type code: str
        """
        tmp_activities = activities.get('activites', None)
        if not tmp_activities:
            return
        acti = [activity for activity in tmp_activities
            if activity.get('type_code') == 'proj'
            and activity.get('project_title') != None]
        if len(list(acti)) > 0:
            self.exporter.add_event_group(code, title)
        for activity in acti:
            self.add_project(
                activity['begin'][:10], activity['end'][:10],
                str(activity['title']), code)

from Exporters import Excel, Calendar
import YAWAEI.YAWAIE.intranet as YAWAEI
import sys

if __name__ == '__main__':
    _type = sys.argv[1]
    intra = YAWAEI.AutologinIntranet(sys.argv[2])
    title = sys.argv[3]
    if len(sys.argv) > 5:
        pedago = sys.argv[4]
    else:
        pedago = ""
    if _type == 'ics':
        i = Calendar(title)
    elif _type == 'xlsx':
        i = Excel(title)
    else:
        print("Unrocognized mode type")
        print('Usage:\t\n\tpipenv run python Planning.py <mode> <token> <title>' 
              ' <pedago>\n<mode> :\n\t- ics\n\t- xlsx\n<token> : intranet autol'
              'ogin token\n<title> : name of the planning\n<pedago> : email '
              'address of the filtered pedago')
        sys.exit(False)
    p = Planning(i, intra, ('08', '11'), [pedago])
