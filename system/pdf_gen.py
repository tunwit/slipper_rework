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
from openpyxl.styles import Font
import pandas as pd
from pathlib import Path
from jinja2 import Environment, FileSystemLoader
import asyncio
from pyppeteer import launch

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
        self.storage_dir = self.get_storage_dir()
        self.slip_dir = self.get_slip_dir()
        self.dfs = self.ex_to_df()
        self.people = None
        self.call_back = None
        self.complete = False

    def ex_to_df(self):
        dfs = pd.read_excel(self.path,header=1,sheet_name=None)
        return self.clean_sheet(dfs)
    
    def clean_sheet(self,dfs:dict):
        to_pop = [df for df in dfs if not df.startswith("D")]
        for b in to_pop:
            dfs.pop(b)

        cleaned_dfs = {}
        for sheet_name, df in dfs.items():
            df.columns.values[0] = "รหัสพนักงาน"

            if "รหัสพนักงาน" in df.columns and "ชื่อ-นามสกุล" in df.columns:
                first_col = df.columns[0]
                df[first_col] = df[first_col].astype(pd.Int64Dtype()).astype(str)
                df.fillna(0, inplace=True)
                print(df.head())
                new_name = sheet_name.replace("D", "", 1)  # remove only the first "D"
                cleaned_dfs[new_name] = df.dropna(subset=["รหัสพนักงาน", "ชื่อ-นามสกุล"])


        return cleaned_dfs
    
    def get_storage_dir(self):
        output = Path(f"{os.getcwd()}\storage\{SHOP_NAME}")
        return output
    
    def get_slip_dir(self):
        output = Path(f"{os.getcwd()}\slip\{SHOP_NAME}")
        return output
    
    def re_init(self):
        self.people = None
        self.complete = False

    def get_lang(self,lang):
        try:
            with open(f"data\\languages\\{lang}.json", "r", encoding='utf-8') as json_file:
                data = json.load(json_file)
        except FileNotFoundError:
            with open(f"data\\languages\\th.json", "r", encoding='utf-8') as json_file:
                data = json.load(json_file)
        return data
    
    def get_value(self,employee,col:str):
        result = employee[col]
        if result == None or result == '0':
            result = "-"
        return result
    
    def get_round(self,branch):
        return self.dfs[branch].shape[0]

    
    def get_all_round(self):
        i=0
        for sheet in self.sources:
            salib = self.get_round(sheet)
            i += salib
        return i

    def mkdir(self,branch):
            if not os.path.exists(os.path.join(self.slip_dir,branch)):
                os.makedirs(os.path.join(self.slip_dir,branch))
            if not os.path.exists(os.path.join(self.storage_dir,branch)):
                os.makedirs(os.path.join(self.storage_dir,branch))

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


    def format_value(self,value, fmt):
        if fmt == "float":
            return f"{value:,.2f}"
        elif fmt == "int":
            return f"{int(value)}"
        else:
            return f"{value}"
    
    def html_to_pdf_sync(self,html_path, pdf_path):
        async def main():
            browser = await launch(headless=True,
                        executablePath=r"C:\Program Files\Google\Chrome\Application\chrome.exe",handleSIGINT=False,
        handleSIGTERM=False,
        handleSIGHUP=False,
        dumpio=False, 
                        args=['--no-sandbox'])
            page = await browser.newPage()
            await page.goto(r'file:///C:\Users\Admin\Desktop\wit\slipper_rework\out.html', waitUntil='networkidle0')
            await page.pdf({
                'path': pdf_path,
                'format': 'A4',
                'landscape': True,
                'printBackground': True,
                'preferCSSPageSize': True,
                'margin': {'top': '10mm', 'bottom': '10mm', 'left': '10mm', 'right': '10mm'},
            })
            await browser.close()

        asyncio.run(main())

    def build_section(self,employee,employee_data,config,date_m,t):
        earnings = [{"label" : t(field['label_key']),
                             "value":employee_data[field['key']],
                             "unit": t(field["unit_key"]),
                             "formatted": self.format_value(employee_data[field['key']], field['format_key'])}
                             for field in config['earnings']['fields'] if field['display']]
                
        deduction = [{"label" : t(field['label_key']),
                        "value":employee_data[field['key']],
                        "unit": t(field["unit_key"]),
                        "formatted": self.format_value(employee_data[field['key']], field['format_key'])} 
                        for field in config['deduction']['fields'] if field['display']]
        
        details = [{"label" : t(field['label_key']),
                        "value":employee_data[field['key']],
                        "unit": t(field["unit_key"]),
                        "formatted": self.format_value(employee_data[field['key']], field['format_key'])} 
                        for field in config['details']['fields'] if field['display']]
        
        net = {"label":t(config['total']['label_key']),
                "value":employee_data[config['total']['key']],
                "unit": t(config['total']["unit_key"]),
                "formatted": self.format_value(employee_data[config['total']['key']], config['total']['format_key'])}

        total_earnings = {
            "label":t('totale'),
            "value":sum(item["value"] for item in earnings),
            "unit": t('฿'),
            "formatted": self.format_value(sum(item["value"] for item in earnings),"float")}
        

        total_deduction = {
            "label":t('totled'),
            "value":sum(item["value"] for item in deduction),
            "unit": t('฿'),
            "formatted": self.format_value(sum(item["value"] for item in deduction),"float")
        }

        context = {
        "company":config['company'],
        "employee": employee,
        "sections": {
            "earnings": earnings,
            "deduction": deduction,
            "details":details,
            "net":net
        },
        "totals": {
                "earnings": total_earnings,
            "deduction": total_deduction,
        },
        "payPeriod":date_m.strftime("%B %Y"),
        "timeStamp":datetime.now().strftime("%d %B %Y %H:%M:%S")
        }
        return context
        
    def extract_convert(self,data:dict,date_m:date):
        print(self.dfs.keys())
        global creating
        creating = True

        self.re_init()
        env = Environment(loader=FileSystemLoader('data/template'))
        with open("config_2.json","r",encoding="utf8") as config:
            config = json.load(config)['haris_slip_details']

        template = env.get_template(config['template'])
        for branch in data.keys():
            self.mkdir(branch)
            for _id in data[branch]:
                employee_data = self.dfs[branch].loc[self.dfs[branch]["รหัสพนักงาน"] == _id].squeeze()

                def t(key): 
                    l = self.get_lang(employee_data['ภาษา'].strip())
                    return l.get(key,key)
                
                employee = {
                    "name" : employee_data["ชื่อ-นามสกุล"],
                    "id": employee_data["รหัสพนักงาน"],
                    "position": employee_data["ตำแหน่ง"],
                    "branch":branch
                }
                
                context = self.build_section(employee,employee_data,config,date_m,t)
               
                html_out = template.render(context=context,t=t)
                filename = f'{employee["id"]}_{employee['name']}'
                path_strorage = os.path.join(self.storage_dir,branch,filename)

                #create pay slip html
                with open(path_strorage+".html",'w', encoding='utf-8') as f:
                    f.write(html_out)
                
                #create pay slip pdf
                self.html_to_pdf_sync(path_strorage+".html",path_strorage+".pdf")

                path_slip = os.path.join(self.slip_dir,branch,filename)
                shutil.copy2(path_strorage+".pdf", path_slip+".pdf")

        self.complete = True
        self.progress()
        creating = False

        # for df in self.dfs:
        #     for i in self.temporary.sheetnames:
        #         if i != "สลิป":
        #             self.temporary.remove(self.temporary[i])
        #     for i in range(self.get_round(sheet)):
        #             sheet_title = sheet.title.replace('D','')
        #             i += 3
        #             if not self.get_value(sheet,2,i) in people:
        #                 continue
        #             index += 1
        #             self.mkdir(sheet_title)
        #             self.progress(index,self.get_value(sheet,2,i),sheet_title)
        #             respound = self.get_lang(self.get_value(sheet,27,i))
        #             img = openpyxl.drawing.image.Image(LOGO_PATH)
        #             img.anchor = 'B1'
        #             ws = [i.title for i in self.salib]
        #             salib = self.temporary[ws[0]]
        #             salib.add_image(img)
                    
        #             salib["C1"] = respound["address"][sheet_title]["adline1"]
        #             salib["C2"] = respound["address"][sheet_title]["adline2"]
        #             salib["C3"] = respound["address"][sheet_title]["adline3"]
        #             salib["C4"] = sheet_title #สาขา
        #             salib["B8"] = f"{respound['ofmonth'].format(month=month[date_m.strftime('%B')],year=date_m.year)}"
        #             # "Key":[text_pos,value_pos,col_pos]

        #             email = self.get_value(sheet,SLIP_DETAIL['email_col'],i)
        #             for field in SLIP_DETAIL.items():
        #                 if field[0] == "email_col":
        #                     continue
        #                 key,text_pos ,value_pos ,col_pos = field[0],field[1][0],field[1][1],field[1][2]
        #                 salib[text_pos] = respound[key]
        #                 if value_pos:
        #                     if type(col_pos) == int:
        #                         salib[value_pos] = self.get_value(sheet,col_pos,i)
        #                     else:
        #                         salib[value_pos] = col_pos
                    
        #             filename = f"{self.get_value(sheet,2,i)},{email},0,{date_m.strftime('%B')},{datetime.now().strftime('%d%m%y%H%M%S')}"
        #             print(filename)
        #             finalpath_ex = os.path.join(self.output_dir,sheet_title,f"{filename}.xlsx")
        #             self.temporary.save(finalpath_ex)
        #             # wb = app.Workbooks.Open(finalpath_ex)
        #             # wb.ActiveSheet.PageSetup.Orientation = 2
        #             # wb.ActiveSheet.PageSetup.Zoom = False
        #             # wb.ActiveSheet.PageSetup.FitToPagesTall = 1
        #             # wb.ActiveSheet.PageSetup.FitToPagesWide = 1
        #             # wb.ActiveSheet.ExportAsFixedFormat(0,os.path.join(self.output_dir,sheet_title,f"{filename}"))
        #             # wb.Save()
        #             # wb.Close()   
        #             os.remove(finalpath_ex)
        #     shutil.copyfile(self.path,self.temporaries)
        #     self.temporary = load_workbook(self.temporaries,data_only=True)
        #     time.sleep(0.1)
        # os.remove(self.temporaries)
        # self.complete = True
        # self.progress()
        # creating = False