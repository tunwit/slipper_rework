
import os
import ssl
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import tempfile
from email.mime.image import MIMEImage
from email.mime.application import MIMEApplication
import pypdfium2 as pdfium
from email.utils import formatdate
import time
import threading
from datetime import datetime
from setup_config import EMAIL_ATTEMP ,SENDER ,PASSWORD, FROM_EMAIL

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
            Createat_converted:{datetime.strptime(person['createat'],'%d%m%y%H%M%S')}
            """
        msg.attach(MIMEText(body,'plain'))
        return msg

    def msg_production_gen(self,person):  
        msg = MIMEMultipart()
        msg['From'] = FROM_EMAIL
        msg['To'] = person['email']
        msg['subject'] = f'สลิปเงินเดือนของ {person["name"]}'
        msg['Date'] = formatdate(localtime=True)
        
        body = f"""
                เรียนคุณ {person['name']} 
        นี่คือใบเเจ้งเงินเดือนประจำเดือน {person['ofmonth']} หากมีปัญหาหรือขอผิดพลาดประการใดกรุณาติดต่อผู้ดูเเล"""

        # msg.attach(MIMEText(body)) bug
        msg.attach(MIMEText('<img src="cid:image1" width="1000" height="772">', 'html'))
        
        pdf = pdfium.PdfDocument(person['path'])
        page = pdf.get_page(0)
        pil_image = page.render(scale = 300/72).to_pil()
        pdf.close()

        temp_file = tempfile.NamedTemporaryFile(delete=False,suffix='.png')
        pil_image.save(temp_file.name)
        image_data = temp_file.read()
        temp_file.close()

        img = MIMEImage(image_data,_subtype="png")
        img.add_header('Content-ID', '<image1>')
        msg.attach(img)

        with open(person["path"],'rb') as f:   
            attach = MIMEApplication(f.read(),_subtype="pdf")
        attach.add_header('Content-Disposition','attachment',filename=f"เงินเดือนของ {person['name']} ประจำเดือน {month[person['ofmonth']]} {datetime.strptime(person['createat'],'%d%m%y%H%M%S').year+543}.pdf")
        msg.attach(attach)
        return msg

    def send_emails(self,person,length):
            self.index += 1
            if person['email'] == "-":
                return
            context = ssl.create_default_context()
            with smtplib.SMTP_SSL('smtp.gmail.com',465,context = context) as smtp:
                smtp.login(SENDER,PASSWORD)            
                msg = self.msg_production_gen(person)
                success = False
                attemp = 0
                while not success and attemp < EMAIL_ATTEMP:
                    attemp += 1
                    try:
                        smtp.sendmail(SENDER,person['email'],msg.as_string())
                        success = True
                    except Exception as e:
                        print(f"Fail to send mail to {person['name']} | {person['email']} trying {attemp}/{EMAIL_ATTEMP} due to {e}")
                        time.sleep(0.4)

                self.progress(self.index,person['name'],person['branch'])
                if not success:
                    print(f"Unable to send mail to {person['name']} | {person['email']}")
                    return
                smtp.quit()
                print(f'Mail has send to {person["name"]} | {person["email"]} {self.index}/{length}')

                path_name:str = os.path.split(person["path"])
                checker = '1'
                new_name = ','.join([person['name'],person["email"],checker,person["ofmonth"],person["createat"]])+'.pdf'
                new_path = os.path.join(path_name[0],new_name)
                try:
                    os.rename(person["path"],new_path)
                except:pass

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
         