from win32com import client
from unittest import mock
# Set max font family value to 100
p = mock.patch('openpyxl.styles.fonts.Font.family.max', new=100)
p.start()

import os
import json
import time
import time
from datetime import date
from datetime import datetime
import openpyxl
from openpyxl import load_workbook
import shutil
from setup_config import SLIP_DETAIL, LOGO_PATH, SHOP_NAME
import tempfile
import threading

creating = False

month = {
        "January":"มกราคม", 
        "February":"กุมภาพันธ์", 
        "March":"มีนาคม", 
        "April":"เมษายน", 
        "May":"พฤษภาคม", 
        "June":"มิถุนายน", 
        "July":"กรกฎาคม", 
        "August":"สิงหาคม", 
        "September":"กันยายน",
        "October":"ตุลาคม", 
        "November":"พฤศจิกายน", 
        "December":"ธันวาคม"
    }


class excel():
    def __init__(self,path) -> None:
        self.path = path
        self.output_dir = self.get_output_dir()
        self.sources = self.get_sources()["sources"]
        self.salib = self.get_sources()["salib"]
        self.people = None
        self.call_back = None
        self.complete = False
        self.index = 0

    def get_output_dir(self):
        output = f"{os.getcwd()}\slip\{SHOP_NAME}"
        return output

    def re_init(self):
        self.people = None
        self.complete = False

    def get_temporaries(self):
        temporaries = 'temporary.xlsx'
        if os.path.exists(f"{temporaries}"):
            os.remove(f"{temporaries}")
        return temporaries
    
    def get_sources(self):
        data = load_workbook(self.path,data_only=True)
        salib = [i for i in data if "สลิป" in i.title and "Data" not in i.title]
        sources = [ i for i in data if i.title.startswith('D') and "Data" not in i.title]
        return {
            "sources":sources,
            "salib":salib
            }
    
    def get_lang(self,lang):
        try:
            with open(f"data\\languages\\{lang}.json", "r", encoding='utf-8') as json_file:
                data = json.load(json_file)
        except FileNotFoundError:
            with open(f"data\\languages\\th.json", "r", encoding='utf-8') as json_file:
                data = json.load(json_file)
        return data
    
    def get_value(self,source,col,i):
        result = source.cell(row=i, column=col).value
        if result == None or result == '0':
            result = "-"
        return result
    
    def get_round(self,source):
        i=0
        while source.cell(row=i+3, column=1).value != None or source.cell(row=i+3, column=2).value != None:
            i+=1
        return i
    
    def get_all_round(self):
        i=0
        for sheet in self.sources:
            salib = self.get_round(sheet)
            i += salib
        return i

    def mkdir(self,branch):
            if not os.path.exists(os.path.join(self.output_dir,branch)):
                os.makedirs(os.path.join(self.output_dir,branch))

    def progress(self,index=None,current=None,branch=None):
        if self.call_back:
            data = {
                'status':'complete'if self.complete else 'processing',
                'all':len(self.people),
            }
            if current:
                data.update({'current':current})
            if index:
                data.update({'index':index,
                             'percentage':(index*100)/len(self.people)})
            if index:
                data.update({'branch':branch})
            self.call_back(data)

    def proceed(self,sheet,i,date_m):
        with tempfile.NamedTemporaryFile(suffix=".xlsx",delete=False) as temp_excel :
            print(temp_excel.name)
            shutil.copyfile(self.path,temp_excel.name)
            copy = load_workbook(temp_excel)
            sheet_title = sheet.title.replace('D','')
            self.index += 1
            self.mkdir(sheet_title)
            self.progress(self.index,self.get_value(sheet,2,i),sheet_title)
            respound = self.get_lang(self.get_value(sheet,27,i))
            img = openpyxl.drawing.image.Image(LOGO_PATH)
            img.anchor = 'B1'
            salib = copy['สลิป']

            for sheetname in copy.sheetnames: # delete all unnecessary sheet EXCEPT Slip
                if 'สลิป' not in sheetname:
                    copy.remove(copy[sheetname])

            salib.add_image(img)
            salib["C1"] = respound["address"][sheet_title]["adline1"]
            salib["C2"] = respound["address"][sheet_title]["adline2"]
            salib["C3"] = respound["address"][sheet_title]["adline3"]
            salib["C4"] = sheet_title #สาขา
            salib["B8"] = f"{respound['ofmonth'].format(month=month[date_m.strftime('%B')],year=date_m.year)}"
            # "Key":[text_pos,value_pos,col_pos]

            for field in SLIP_DETAIL.items():
                key,text_pos ,value_pos ,col_pos = field[0],field[1][0],field[1][1],field[1][2]
                salib[text_pos] = respound[key]
                if value_pos:
                    if type(col_pos) == int:
                        salib[value_pos] = self.get_value(sheet,col_pos,i)
                    else:
                        salib[value_pos] = col_pos
            copy.save(temp_excel.name)

            app = client.DispatchEx("Excel.Application")
            app.Interactive = False
            app.Visible = False
            wb = app.Workbooks.Open(temp_excel.name)
            wb.ActiveSheet.PageSetup.Orientation = 2
            wb.ActiveSheet.PageSetup.Zoom = False
            wb.ActiveSheet.PageSetup.FitToPagesTall = 1
            wb.ActiveSheet.PageSetup.FitToPagesWide = 1
            
            filename = f"{self.get_value(sheet,2,i)},{self.get_value(sheet,22,i)},0,{date_m.strftime('%B')},{datetime.now().strftime('%d%m%y%H%M%S')}"
            wb.ActiveSheet.ExportAsFixedFormat(0,os.path.join(self.output_dir,sheet_title,f"{filename}"))
            temp_excel.close()

    def extract_convert(self,people_pre:list,date_m:date):
        global creating
        creating = True
        people= []
        for person in people_pre:
            person:str
            new = person.replace('[/size]','')
            people.append(new)
        self.re_init()
        self.people = people

        tasks = []
        for sheet in self.sources:
            for i in range(self.get_round(sheet)):
                i += 3
                if not self.get_value(sheet,2,i) in people:
                    continue
                task = threading.Thread(target=self.proceed,args=(sheet,i,date_m,))
                tasks.append(task)
                task.start()

        for task in tasks:
            task.join()  

        self.complete = True
        self.progress()
        creating = False