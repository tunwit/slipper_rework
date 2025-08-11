import os
import json
from pathlib import Path
from datetime import date
from datetime import datetime
import shutil
from setup_config import SLIP_DETAIL, PDF_GENERATOR_CONCURRENCY, SHOP_NAME
import pandas as pd
from pathlib import Path
from jinja2 import Environment, FileSystemLoader
import asyncio
from pyppeteer import launch
import logging
import sys
import contextlib
import time


logging.getLogger('pyppeteer').setLevel(logging.WARNING)
logging.getLogger('websockets').setLevel(logging.ERROR)

@contextlib.contextmanager
def suppress_output():
    with open(os.devnull, 'w') as devnull:
        old_stdout = sys.stdout
        old_stderr = sys.stderr
        sys.stdout = devnull
        sys.stderr = devnull
        try:
            yield
        finally:
            sys.stdout = old_stdout
            sys.stderr = old_stderr

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
        self.call_back = None
        self.complete = False
        self.total_rows = 0
        self.jobs = []
        self.progress_percent = 0
        self.pages = []
        self.concurrent_count = PDF_GENERATOR_CONCURRENCY * 2
        self.semaphore = asyncio.Semaphore(self.concurrent_count)
    
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
                new_name = sheet_name.replace("D", "", 1)  # remove only the first "D"
                cleaned_dfs[new_name] = df.dropna(subset=["รหัสพนักงาน", "ชื่อ-นามสกุล"])
                first_col = cleaned_dfs[new_name].columns[0]
                cleaned_dfs[new_name][first_col] = cleaned_dfs[new_name][first_col].astype(pd.Int64Dtype()).astype(str)
                cleaned_dfs[new_name].fillna(0, inplace=True)
        
        return cleaned_dfs
    
    def get_storage_dir(self):
        output = Path.cwd() / "storage" / SHOP_NAME

        return output
    
    def get_slip_dir(self):
        output = Path.cwd() / "slip" / SHOP_NAME
        return output
    
    def re_init(self):
        self.total_rows = 0
        self.complete = False

    def get_lang(self,lang):
        path = Path.cwd() / "data" / "languages"
        try:
            with open(path / f"{lang}.json", "r", encoding='utf-8') as json_file:
                data = json.load(json_file)
        except FileNotFoundError:
            with open(path / "th.json", "r", encoding='utf-8') as json_file:
                data = json.load(json_file)
        return data
    
    def get_value(self,employee,col:str):
        result = employee[col]
        if result == None or result == '0':
            result = "-"
        return result

    def get_round(self,branch):
        return self.dfs[branch].shape[0]
    

    def mkdir(self,branch):
            if not os.path.exists(self.slip_dir / branch):
                os.makedirs(self.slip_dir / branch)
            if not os.path.exists(self.storage_dir / branch):
                os.makedirs(self.storage_dir / branch)

    def progress(self,current=None,branch=None,file='html'):
        self.progress_percent+=1
        if self.call_back:
            data = {
                'status':'complete'if self.complete else 'processing',
                'file':file,
                'all':self.total_rows*2,
            }
            if current:
                data.update({'current':current})
            data.update({'percentage':(self.progress_percent*100)/(self.total_rows*2)})
            if branch:
                data.update({'branch':branch})
            self.call_back(data)


    def format_value(self,value, fmt):
        if fmt == "float":
            return f"{value:,.2f}"
        elif fmt == "int":
            return f"{int(value)}"
        else:
            return f"{value}"
    
    def render_to_html(self,context,t,template,path_storage,filename) -> str:
        html = template.render(context=context,t=t)

        final_path = path_storage / filename
        with open(final_path.with_suffix(".html"),'w', encoding='utf-8') as f:
            f.write(html)

        return html

    async def render_to_pdf(self, html_content, pdf_path, path_storage, context):
        async with self.semaphore:
            page = self.pages.pop()
            try:
                await page.setContent(html_content)
                await page.pdf({
                    'path': pdf_path,
                    'format': 'A4',
                    'landscape': False,
                    'printBackground': True,
                    'preferCSSPageSize': True,
                    'margin': {'top': '10mm', 'bottom': '10mm', 'left': '10mm', 'right': '10mm'},
                })
            finally:
                self.pages.append(page)
            
        #create pay slip pdf
        file_name = path_storage / "pay_slip_pdf"
        start_time = time.perf_counter()
        self.createMetaData(path_storage,context)
        end_time = time.perf_counter()
        metadata_io_time = end_time - start_time
        print(f"Metadata creation took {metadata_io_time:.17f} seconds")
        key = f'{context['employee']["id"]}_{context['employee']['name']}'.strip()
        path_slip = self.slip_dir / context['employee']['branch'] / key
        shutil.copy2(file_name.with_suffix(".pdf"), path_slip.with_suffix(".pdf"))
        self.progress(current=context['employee']["name"],branch=context['employee']['branch'],file='pdf')

    async def init_pages(self, browser, count=5):
        for _ in range(count):
            page = await browser.newPage()
            self.pages.append(page)

    async def html_to_pdf_batch(self):
            if os.name == 'nt':
                exe = Path("C:\Program Files\Google\Chrome\Application\chrome.exe")
            else:
                exe = Path(r"/Applications/Google Chrome.app/Contents/MacOS/Google Chrome")
            with suppress_output():
                browser = await launch(headless=True,
                            executablePath=exe,handleSIGINT=False,
                            handleSIGTERM=False,
                            handleSIGHUP=False,
                            dumpio=False, 
                            args=['--no-sandbox',
                            '--disable-setuid-sandbox',
                            '--disable-dev-shm-usage',
                            '--disable-extensions',
                            '--disable-background-networking',
                            '--disable-sync',
                            '--disable-translate',
                            '--hide-scrollbars',
                            '--metrics-recording-only',
                            '--mute-audio',
                            '--no-first-run',
                            '--safebrowsing-disable-auto-update',
                            '--disable-background-timer-throttling',
                            '--disable-renderer-backgrounding',
                            '--disable-backgrounding-occluded-windows',
                            '--memory-pressure-off'])
                
            env = Environment(loader=FileSystemLoader('data/template'))
            template_pdf = env.get_template(SLIP_DETAIL['template_pdf'])
            template_email = env.get_template(SLIP_DETAIL['template_email'])
            lang_th = self.get_lang('th')
            lang_en = self.get_lang('en')
            
            tasks = []
            await self.init_pages(browser, count=self.concurrent_count)

            for path_storage, context in self.jobs:
                if context['employee']['locale'] == 'th':
                    lang = lang_th
                else:
                    lang = lang_en

                def t(key):
                    return lang.get(key, key)
                
                self.render_to_html(context, t, template_email, path_storage, "pay_slip_email")
                html_content_pdf = self.render_to_html(context, t, template_pdf, path_storage, "pay_slip_pdf")
                self.progress(current=context["employee"]["name"],branch=context["employee"]["branch"],file='html')

                pdf_path = path_storage / "pay_slip_pdf.pdf"
                task = self.render_to_pdf(html_content_pdf, pdf_path, path_storage, context)
                tasks.append(task)

            await asyncio.gather(*tasks)

            for page in self.pages:
                await page.close()

            await browser.close()


    def createMetaData(self,storage_path,context):
        json_file = storage_path / "metadata.json"
        meta_data = {
            "employee_id": context['employee']['id'],
            "employee_name": context['employee']['name'],
            "email": context['employee']['email'],
            "branch":context['employee']['branch'],
            "pay_period": context['payPeriod'],
            "pdf_path": str((storage_path / "pay_slip_pdf").with_suffix(".pdf")),
            "html_path": str((storage_path / "pay_slip_pdf").with_suffix(".html")),
            "html_email_path": str((storage_path / "pay_slip_email").with_suffix(".html")),
            "meta_data_path":str(json_file),
            "mail_sent": False,
            "locale": context['locale'],
            "created_at": context['timeStamp']
        }
        
        with open(json_file,'w',encoding="utf-8") as f:
            json.dump(meta_data, f, ensure_ascii=False, indent=2)
        

    def build_section(self,employee,employee_data,date_m,t):
        earnings = [{"label" : t(field['label_key']),
                             "value":employee_data[field['key']],
                             "unit": t(field["unit_key"]),
                             "formatted": self.format_value(employee_data[field['key']], field['format_key'])}
                             for field in SLIP_DETAIL['earnings']['fields'] if field['display']]
                
        deduction = [{"label" : t(field['label_key']),
                        "value":employee_data[field['key']],
                        "unit": t(field["unit_key"]),
                        "formatted": self.format_value(employee_data[field['key']], field['format_key'])} 
                        for field in SLIP_DETAIL['deduction']['fields'] if field['display']]
        
        details = [{"label" : t(field['label_key']),
                        "value":employee_data[field['key']],
                        "unit": t(field["unit_key"]),
                        "formatted": self.format_value(employee_data[field['key']], field['format_key'])} 
                        for field in SLIP_DETAIL['details']['fields'] if field['display']]
        
        net = {"label":t(SLIP_DETAIL['total']['label_key']),
                "value":employee_data[SLIP_DETAIL['total']['key']],
                "unit": t(SLIP_DETAIL['total']["unit_key"]),
                "formatted": self.format_value(employee_data[SLIP_DETAIL['total']['key']], SLIP_DETAIL['total']['format_key'])}

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
        "company":{
            "name":SLIP_DETAIL["company"]["name"],
            "branch":SLIP_DETAIL["company"]["branch"][employee["branch"]]
        },
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
        "locale":employee_data['ภาษา'],
        "timeStamp":datetime.now().strftime("%d %B %Y %H:%M:%S")
        }
        return context
    
    def extract_convert(self,data:dict,date_m:date):
        global creating
        creating = True

        self.re_init()
        self.total_rows = sum(len(df) for df in data.values())

        #pre load langauge
        lang_th = self.get_lang('th')
        lang_en = self.get_lang('en')

        for branch in data.keys():
            self.mkdir(branch)
            df_map = {row["รหัสพนักงาน"]: row for _, row in self.dfs[branch].iterrows()}
            for _id in data[branch]:
                employee_data = df_map[_id]

                if employee_data['ภาษา'] == 'th':
                    lang = lang_th
                else:
                    lang = lang_en

                def t(key):
                    return lang.get(key, key)
                
                employee = {
                    "name" : employee_data["ชื่อ-นามสกุล"],
                    "id": employee_data["รหัสพนักงาน"],
                    "position": employee_data["ตำแหน่ง"],
                    "email":employee_data['Email'],
                    "locale": employee_data['ภาษา'],
                    "branch":branch
                }
                
                context = self.build_section(employee,employee_data,date_m,t)

                key = f'{employee["id"]}_{employee["name"]}'.strip()
                path_storage = self.storage_dir / branch / key

                #create pay slip html
                if not os.path.exists(path_storage):
                    os.makedirs(path_storage, exist_ok=True)
                
                self.jobs.append((path_storage,context))

        asyncio.run(self.html_to_pdf_batch())
        self.complete = True
        self.progress()
        creating = False
        self.progress_percent = 0