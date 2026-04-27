
import os
import ssl
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import tempfile
from email.mime.image import MIMEImage
from email.mime.application import MIMEApplication
import resend
from email.utils import formatdate
import time
import threading
from datetime import datetime
from setup_config import EMAIL_ATTEMP ,SENDER ,RESEND_API, FROM_EMAIL
import json
import base64

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


class send_email():
    
    def __init__(self) -> None:
        self.people = None
        self.call_back = None
        self.complete = False
        self.index = 0
        resend.api_key = RESEND_API
    
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

    def msg_test_gen(self,person):
        msg = MIMEMultipart()
        msg['From'] = FROM_EMAIL
        msg['To'] = person['email']
        msg['subject'] = f'ทดสอบระบบ System Testing'
        msg['Date'] = formatdate(localtime=True)
        body = f"""
            Testing System
            กำลังทดสอบระบบ...
            
            Name:{person['name']}
            Path:{person['path']}
            Email:{person['email']}
            Branch:{person['branch']}
            File_name:{person['file_name']}
            Ofmonth:{person['ofmonth']}
            Createat_raw:{person['createat']}
            Createat_converted:{datetime.fromisoformat(person['createat'])}
            """
        msg.attach(MIMEText(body,'plain'))
        return msg

    def msg_production_gen(self,person):  
        with open(person['html_email_path'], "r", encoding="utf-8") as f:
            html_content = f.read()

        with open(person["pdf_path"],'rb') as f:   
            attach = f.read()

        params: resend.Emails.SendParams = {
                "from": f"{FROM_EMAIL} <{SENDER}>",
                "to": [person['email']],
                "subject":  f'สลิปเงินเดือนของ {person["employee_name"]}',
                "html": html_content,
                "attachments":[{
                   "filename":f"เงินเดือนของ {person['employee_name']} ประจำเดือน {person['pay_period']}.pdf",
                   "content":base64.b64encode(attach).decode()
                }]
        }
        return params

    def send_emails(self,person,length):
            self.index += 1
            if person['email'] == "-" or person['email'] == 0 or person['email'] == "":
                return
            params = self.msg_production_gen(person)
            success = False
            attemp = 0
            while not success and attemp < EMAIL_ATTEMP:
                attemp += 1
                try:
                    email = resend.Emails.send(params)
                    success = True
                except Exception as e:
                    print(f"Fail to send mail to {person['employee_name']} | {person['email']} trying {attemp}/{EMAIL_ATTEMP} due to {e}")
                    time.sleep(0.4)

            self.progress(self.index,person['employee_name'],person['branch'])
            if not success:
                print(f"Unable to send mail to {person['employee_name']} | {person['email']}")
                return
            print(f'Mail has send to {person["employee_name"]} | {person["email"]} {self.index}/{length}')
            
            person['mail_sent'] = True
            with open(person["meta_data_path"],'w',encoding="utf-8") as f:
                json.dump(person, f, ensure_ascii=False, indent=2)

    def send(self,people):

        self.people = people

        # text = """
        # เรียน {name},\n\n
        # \tใบสลิปเงินเดือนของ {name} ประจำเดือน {DY} ของบริษัท ไอแอมฟู้ด จำกัด\n
        # หากมีข้อผิดพลาดประการใดขออภัยไว้ ณ ที่นี้\n\n
        # \t\t\t\t\tจึงเรียนมาเพื่อทราบ\n
        # \t\t\t\t\tบริษัท ไอแอมฟู้ด จำกัด
        # """
        threads = []
        length = len(people)
        for i,person in enumerate(people):
            thread = threading.Thread(target=self.send_emails, args=(person,length,))
            threads.append(thread)
            thread.start()
            time.sleep(0.4)
            if i %5 == 0:
                time.sleep(2)
                

        for thread in threads:
            thread.join()

        self.complete = True
        self.progress()
         