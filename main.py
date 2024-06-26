from kivy.core.text import LabelBase
from kivymd.uix.button import MDRectangleFlatButton
from kivymd.uix.label import MDLabel
from kivymd.uix.boxlayout import BoxLayout
from kivymd.uix.screen import MDScreen
from kivymd.uix.screenmanager import MDScreenManager
from kivymd.uix.boxlayout import MDBoxLayout
from kivymd.uix.list import IconRightWidget
from kivymd.uix.list import IconLeftWidget
from kivymd.app import MDApp
from kivy.lang import Builder
from kivymd.uix.selectioncontrol import MDCheckbox 
from kivymd.uix.card import MDSeparator 
from kivymd.uix.scrollview import MDScrollView 
from kivymd.uix.list import MDList ,IRightBodyTouch
from kivymd.uix.button import MDRectangleFlatButton 
from kivymd.uix.dialog import MDDialog
from kivymd.uix.label import MDIcon
from kivymd.uix.button import MDFlatButton
from kivymd.uix.button import MDRaisedButton
from kivymd.uix.snackbar import MDSnackbar
from kivymd.uix.progressbar import MDProgressBar
from kivymd.uix.textfield import MDTextField
from kivymd.uix.pickers import MDDatePicker
from kivy.clock import Clock
import os
import tkinter
from tkinter import filedialog
from kivymd.uix.list import OneLineAvatarIconListItem
from kivymd.font_definitions import theme_font_styles
import os
from kivy.properties import  StringProperty,BooleanProperty,ListProperty
from kivymd.uix.expansionpanel import MDExpansionPanel, MDExpansionPanelTwoLine
import os
from datetime import date
from datetime import datetime
import sys
import threading
import json
import re
from setup_config import SHOP_ID,SHOP_NAME,TITLE
from system.send_mail import send_email
from system.pdf_gen import excel 
from system.pdf_gen import creating

"""

------------------- shop_id -------------------

    1   :   haris
    2   :   tukkae

-----------------------------------------------

"""

excel_object = False
gmail_interval = None
employee_interval = None
mouth = {
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

class Content(BoxLayout):
    pass

class list_container(IRightBodyTouch, MDBoxLayout):
    adaptive_height=True

class SlipMaker(MDScreen):
    def on_start(self):
        pass

    excel_path = StringProperty("Please select file")
    going_to_make_slip = []
    excel_object:excel = None
    total_individual_checkbox = 0

    def call_back_create_slip(self,data):
        print(data)
        if data['status'] == 'processing':
            Clock.schedule_once(lambda dt: self.update_progress_slip_bar(data))
        elif data['status'] == 'complete':   
            Clock.schedule_once(lambda dt: self.done_create_slip(),1)

    def done_create_slip(self):
        employee_interval()
        gmail_interval()
        self.progress_dialog.title = "Complete !!"
        self.progress_dialog.dismiss()
        snackbar = MDSnackbar(
                MDLabel(
                text="Dowload Complete !",
                font_name = 'sarabun',
                text_color=self.theme_cls.opposite_bg_normal,
            ),
            pos_hint={"center_x": 0.5},
            size_hint_x=0.5,
            md_bg_color="#ffffff",
            duration = 5
            )
        snackbar.open()

    def update_progress_slip_bar(self, data):
        self.ids['progress_slip_bar'].value = data['percentage']
        self.ids['label_slip'].text = f"{int(data['percentage'])}% creating file for {data['branch']}/{data['current']}"

    def start_creating_slip(self, instance, value:date, date_range):
        self.progress_dialog = MDDialog(
            title=f'[size=28][font=sarabunBold]Creating files[/font][/size]',
            type="custom",
            content_cls=MDBoxLayout(
            orientation="vertical",
            spacing= "30dp",
            padding= "20dp",
            size_hint_y = None,
            height= "50dp",
            ),
            radius=[20, 7, 20, 7]
        )
        self.progress_dialog.auto_dismiss = False
        progress_bar = MDProgressBar(value=0,pos_hint={'center_y':-0.2})
        label = MDLabel(text='Initualizing',theme_text_color='Hint',pos_hint={'center_y':0},font_style='sarabun',font_size= "20dp")
        self.ids['label_slip'] = label
        self.ids['progress_slip_bar'] = progress_bar
        self.progress_dialog.content_cls.add_widget(label)
        self.progress_dialog.content_cls.add_widget(progress_bar)
        self.progress_dialog.open()
        self.excel_object.call_back = self.call_back_create_slip
        thread = threading.Thread(target=self.excel_object.extract_convert, args=(self.going_to_make_slip,value,))
        thread.start()
        gmail_interval.cancel()
        employee_interval.cancel()


    def close_confirm_makeslip_dialog(self,dialog):
        self.confirm_makeslip_dialog.dismiss()

    def select_date(self,dialog):
        self.confirm_makeslip_dialog.dismiss()
        today = datetime.today()
        self.date_dialog = MDDatePicker(
            year=today.year+543,
            month=today.month,
            day=today.day,
            min_year=2400,
            max_year=2666,
            title_input="Input payment date",
            title="Select payment date",
            )
        self.date_dialog.bind(on_save=self.start_creating_slip)
        self.date_dialog.open()

    def create_slip(self,button):
        cancle_button = MDFlatButton(
                        text="CANCEL",
                        on_release=self.close_confirm_makeslip_dialog
                    )
        start_button = MDRaisedButton(
                        text="Next",
                        on_release=self.select_date
                    )
        self.confirm_makeslip_dialog = MDDialog(
            title=f'[size=28][font=sarabunBold]Are you sure to start making slip for {len(self.going_to_make_slip)} people?[/font][/size]',
            text='[size=20][font=sarabun]This process can not be canceled make sure you are ready[/font][/size]',
            type="simple",
            radius=[20, 7, 20, 7],
            buttons=[cancle_button,start_button]
        )
        print(self.going_to_make_slip)
        self.confirm_makeslip_dialog.open()
    
    def individual_selected(self,button:MDCheckbox,pos):
        text = button.parent.parent.parent.text
        pattern = r'\[.*?\]\d+.(.*?)\[/.*?]'
        result = re.search(pattern, text)
        name = result.group(1)
        if button.active:
            self.going_to_make_slip.append(name)
        else:
            try:
                self.going_to_make_slip.remove(name)
            except:pass

        true_state = len(self.going_to_make_slip)
        self.ids['number'].text = str(true_state)
        if true_state > 0:
            self.ids.create_slip.disabled=False
        else:
            self.ids.create_slip.disabled=True

        if true_state == self.total_individual_checkbox:
            self.ids.all_select.active = True
        else:
            self.ids.all_select.active = False


    def all_selected(self,button:MDRaisedButton):
        state = button.active
        for ids in self.ids: 
            if 'checkbox_maker_' in ids:
                self.ids[ids].active = state

    def create_employee_list(self,dt):
        fram1 = MDBoxLayout(
            orientation="vertical",
            md_bg_color= (0.217,0.217,0.217,0.1),
            size_hint= (0.5,1),
            id = 'fram1'
        )
        self.ids['fram1'] = fram1
        fram1.add_widget(
            MDLabel(
                text= "Employee List",
                size_hint= (1,0.1),
                halign= 'center',
                font_style= "sarabunBold",
                font_size= "20sp"
            )
        )
        fram2 = MDBoxLayout(
            orientation="horizontal",
            size_hint= (1,0.08)
        )
        fram2.add_widget(
            MDLabel(
                text= "All",
                size_hint= (0.3,1),
                pos_hint={'center_y':0.75},
                halign= 'right',
                font_style= "sarabun",
                font_size= "25dp"
            )
        )
        all_checkbox = MDCheckbox(
                size_hint= (0.3,1),
                pos_hint={'center_y':0.65},
                halign= 'left',
            ) 
        self.ids["all_select"] = all_checkbox
        all_checkbox.bind(on_press=self.all_selected)
        fram2.add_widget(all_checkbox)
        fram1.add_widget(fram2)
        fram1.add_widget(MDSeparator())
        scroll = MDScrollView()  
        mdlst_slip = MDList()
        self.ids['lst'] = mdlst_slip
        for branch in self.excel_object.sources:
            amount = self.excel_object.get_round(branch)
            content = MDBoxLayout()
            content.adaptive_height = True
            content.orientation = 'vertical'
            self.ids[branch.title+'slip'] = content 
            panel = MDExpansionPanel(
                    icon="office-building-outline",
                    content=content,
                    panel_cls=MDExpansionPanelTwoLine(
                        text=f"[size=25]{branch.title.replace('D','')}[/size]",
                        secondary_text=f"{amount} คน",
                        font_style= "sarabunBold",
                        secondary_font_style= "sarabunBold"
                    )
                )
            mdlst_slip.add_widget(panel)
            for num,rows in enumerate(range(amount),1):
                rows += 3
                item = OneLineAvatarIconListItem(
                    text=f"[size=20]{num}.{self.excel_object.get_value(branch,2,rows)}[/size]",
                    font_style = "sarabun",
                )
                face = IconLeftWidget(
                    icon='account'
                )
                check = list_container()
                checkbox = MDCheckbox()
                checkbox.bind(active=self.individual_selected)
                self.total_individual_checkbox += 1
                self.ids[f"checkbox_maker_{datetime.now().strftime('%f')}"] = checkbox
                checkbox.color_active = self.theme_cls.accent_light
                item.add_widget(face)
                check.add_widget(checkbox)
                item.add_widget(check)
                b = str(branch.title) +'slip'
                self.ids[b].add_widget(item)
        print(self.ids)
        scroll.add_widget(mdlst_slip)
        fram1.add_widget(scroll)
        fram1.add_widget(MDSeparator())
        screen = MDScreen(
            size_hint = (1,0.2)
        )
        create_button = MDRaisedButton(
                size_hint= (None,None),
                text= "Create Slip",
                pos_hint= {'center_x':.6,'center_y':.5},
                disabled = BooleanProperty(True)
            )
        self.ids['create_slip'] = create_button
        create_button.bind(on_press=self.create_slip)
        ico = MDIcon(
            icon='account',
            pos_hint= {'center_x':.1,'center_y':.53},
        )
        number = MDLabel(
            text = "0",
            font_name='sarabun',
            pos_hint= {'center_x':0.7,'center_y':.53}
        )
        self.ids['number'] = number
        screen.add_widget(number)
        screen.add_widget(ico)
        screen.add_widget(create_button)
        fram1.add_widget(screen)
        self.ids.box1.add_widget(fram1)

    def on_browse(self):
        path = self.file_open()
        if path:
            self.excel_object = excel(path = path)
            self.going_to_make_slip = []
            try:
                self.ids.box1.remove_widget(self.ids.fram1)
            except:pass
            Clock.schedule_once(self.create_employee_list)
            self.snackbar = MDSnackbar(
                MDLabel(
                text="Upload Complete !",
                font_name = 'sarabun',
                text_color=self.theme_cls.opposite_bg_normal
            ),
            pos_hint={"center_x": 0.5},
            size_hint_x=0.5,
            md_bg_color="#ffffff"
            )
            self.snackbar.open()

    def file_open(self):
        root = tkinter.Tk()
        root.withdraw()
        # root.iconbitmap("YOUR_IMAGE.ico")
        currdir = os.getcwd()
        tempdir = filedialog.askopenfilename(parent=root, initialdir=currdir, title='Please select excel file',filetypes=[("Excel files",".xlsx .xls")])
        path = None
        if len(tempdir) > 0:
            path = tempdir
            self.excel_path = os.path.basename(tempdir)
        return path
    
class CustomOneLineAvatarIconListItem(OneLineAvatarIconListItem):
    def __init__(self, path,email,name,branch, file_name,*args, **kwargs):
        super().__init__(*args, **kwargs)
        self.path = path
        self.name = name
        self.email = email
        self.branch = branch
        self.file_name = file_name

class GmailSender(MDScreen):
    going_to_send = []
    recent_employee = 0
    total_individual_checkbox = 0
    def on_start(self):
        global gmail_interval
        gmail_interval = Clock.schedule_interval(self.update_lst,0.5)

    def done_send_email(self):
        self.ids['box_email'].clear_widgets()
        Clock.schedule_once(lambda dt: self.add_lst())
        self.progress_dialog.title = "Complete !!"
        self.progress_dialog.dismiss()
        snackbar = MDSnackbar(
                MDLabel(
                text="Send Complete !",
                font_name = 'sarabun',
                text_color=self.theme_cls.opposite_bg_normal,
            ),
            pos_hint={"center_x": 0.5},
            size_hint_x=0.5,
            md_bg_color="#ffffff",
            duration = 5
            )
        snackbar.open()

    def call_back_create_email(self,data):
        if data['status'] == 'processing':
            Clock.schedule_once(lambda dt: self.update_progress_email_bar(data))
        elif data['status'] == 'complete':   
            Clock.schedule_once(lambda dt: self.done_send_email(),1)

    def update_progress_email_bar(self, data):
        self.ids['progress_email_bar'].value = data['percentage']
        self.ids['label_email'].text = f"{int(data['percentage'])}% Sending file for {data['branch']}/{data['current']}"

    def start_send_email(self,dialog):
        self.confirm_sendemail_dialog.dismiss()
        # for path in os.listdir('slip'):
        #     for file in os.listdir('slip/'+path):
        #         name = file.split(',')
        #         name[-1] = '1.pdf'
        #         new_name = ','.join(name)
        #         print(self.going_to_send)
        #         print('----------')
        #         print('slip/'+path+'/'+file)
        #         if 'slip/'+path+'/'+file in self.going_to_send:
        #             os.rename(f'slip/{path}/{file}',f'slip/{path}/{new_name}')
        #             self.ids['box_email'].clear_widgets()
        #             Clock.schedule_once(lambda dt: self.add_lst())
        self.progress_dialog = MDDialog(
            title=f'[size=28][font=sarabunBold]Sending mail[/font][/size]',
            type="custom",
            content_cls=MDBoxLayout(
            orientation="vertical",
            spacing= "30dp",
            padding= "20dp",
            size_hint_y = None,
            height= "50dp"
            ),
            radius=[20, 7, 20, 7]
        )
        self.progress_dialog.auto_dismiss = False
        progress_bar = MDProgressBar(value=0,pos_hint={'center_y':-0.2})
        label = MDLabel(text='Initualizing',theme_text_color='Hint',pos_hint={'center_y':0},font_style='sarabun',font_size= "20dp")
        self.ids['label_email'] = label
        self.ids['progress_email_bar'] = progress_bar
        self.progress_dialog.content_cls.add_widget(label)
        self.progress_dialog.content_cls.add_widget(progress_bar)
        self.progress_dialog.open()
        self.sendemail_object = send_email()
        self.sendemail_object.call_back = self.call_back_create_email
        thread = threading.Thread(target=self.sendemail_object.send, args=(self.going_to_send,))
        thread.start()

    def close_confirm_sendemail_dialog(self,dialog):
        self.confirm_sendemail_dialog.dismiss()

    def send_email_button(self):
        cancle_button = MDFlatButton(
                        text="CANCEL",
                        on_release=self.close_confirm_sendemail_dialog
                    )
        start_button = MDRaisedButton(
                        text="Start",
                        on_release=self.start_send_email
                    )
        self.confirm_sendemail_dialog = MDDialog(
            title=f'[size=28][font=sarabunBold]Are you sure to start send email for {len(self.going_to_send)} people?[/font][/size]',
            text='[size=20][font=sarabun]This process may take a several minutes and can not be cancled[/font][/size]',
            radius=[20, 7, 20, 7],
            buttons=[cancle_button,start_button]
        )
        self.confirm_sendemail_dialog.open()

    def update_lst(self,dt):
        if creating:
            return
        
        f = []
        for path in os.listdir(f'slip/{SHOP_NAME}'):
            for i in os.listdir(f'slip/{SHOP_NAME}/'+path):
                f.append(i)

        if len(f) != self.recent_employee:
            self.ids['box_email'].clear_widgets()
            
            if len(f) != 0:
                Clock.schedule_once(lambda dt: self.add_lst())    
            else:
                Clock.schedule_once(lambda dt: self.no_employee_label())
                self.ids.sendemail_button.disabled=True
                self.ids.sendemail_button.text = f' 0       Send email               '
            self.recent_employee = len(f)

    def individual_selected(self,button:MDCheckbox,pos):
        data = button.parent.parent.parent
        name,email,checker,ofmonth,createat = data.file_name[:-4].split(',')
        payload={
            'path':data.path,
            'name':data.name,
            'email':data.email,
            'branch':data.branch,
            'file_name':data.file_name,
            'ofmonth':ofmonth,
            'createat':createat
        }
        if button.active:
            self.going_to_send.append(payload)
        else:
            try:
                self.going_to_send.remove(payload)
            except:pass

        true_state = len(self.going_to_send)
        self.ids.sendemail_button.text = f' {true_state}       Send email               '

        if true_state > 0:
            self.ids.sendemail_button.disabled=False
        else:
            self.ids.sendemail_button.disabled=True

        if true_state == self.total_individual_checkbox:
            self.ids.all_select_email.active = True
        else:
            self.ids.all_select_email.active = False


    def all_selected(self,button:MDRaisedButton):
        state = button.active
        for ids in self.ids: 
            if 'checkbox_email_' in ids:
                self.ids[ids].active = state

    def add_lst(self):
        self.going_to_send=[]
        self.ids.sendemail_button.text = f' {len(self.going_to_send)}       Send email               '
        fram1 = MDBoxLayout(
            orientation="horizontal",
            size_hint= (1,0.08)
        )
        self.ids['fram1_employee'] = fram1
        fram1.add_widget(
            MDLabel(
                text= "All",
                size_hint= (0.3,1),
                pos_hint={'center_y':0.75},
                halign= 'right',
                font_style= "sarabun",
                font_size= "25dp"
            )
        )
        all_checkbox = MDCheckbox(
            size_hint= (0.3,1),
            pos_hint={'center_y':0.65},
            halign= 'left',
        ) 
        all_checkbox.bind(on_press=self.all_selected)
        self.ids["all_select_email"] = all_checkbox
        fram1.add_widget(all_checkbox)
        self.ids['box_email'].add_widget(fram1)
        scroll = MDScrollView()
        self.ids['scroll_email'] = scroll
        mdlst_employee = MDList()
        scroll.add_widget(mdlst_employee)

        for branch in os.listdir(f'slip/{SHOP_NAME}'):
            content = MDBoxLayout()
            content.adaptive_height = True
            content.orientation = 'vertical'
            self.ids[f"{branch.title}email"] = content 
            panel = MDExpansionPanel(
                icon="office-building-outline",
                content=content,
                pos_hint={'center_y':1},
                panel_cls=MDExpansionPanelTwoLine(
                    text=f"[size=25]{branch}[/size]",
                    secondary_text=f"{len(os.listdir(f'slip/{SHOP_NAME}/'+branch))}คน",
                    font_style= "sarabunBold",
                    secondary_font_style= "sarabunBold"
                )
            )
            #39
            for num,filename in enumerate(os.listdir(f'slip/{SHOP_NAME}/'+branch),1):
                # cleaned_string = regex.sub('\p{M}', '',f"{num}.{filename.split(',')[0]}")
                # name_lenght = len(cleaned_string)
                # less = 39 - name_lenght
                # space = ' '*less)
                try:
                    name,email,checker,ofmonth,createat = filename[:-4].split(',')
                    item = CustomOneLineAvatarIconListItem(
                        text=f"[size=20]{num}.{name}   |   {email}   |   Salary of [b]{ofmonth}[/b]   [color=9C9B9B]Create at {datetime.strptime(createat,'%d%m%y%H%M%S').strftime('%d/%m/%y %H:%M')}[/color][/size]",
                        font_style = "sarabun",
                        path= os.path.join('slip',SHOP_NAME,branch,filename),
                        email=filename.split(',')[1],
                        name=filename.split(',')[0],
                        branch=branch,
                        file_name=filename
                    )
                    if checker == '0':
                        icon = 'account-alert'
                    else:
                        icon = 'account-check'
                    face = IconLeftWidget(
                        icon=icon
                    )
                    check = list_container()
                    checkbox = MDCheckbox()
                    checkbox.bind(active=self.individual_selected)
                    self.total_individual_checkbox += 1
                    self.ids[f"checkbox_email_{datetime.now().strftime('%f')}"] = checkbox
                    checkbox.color_active = self.theme_cls.accent_light
                    item.add_widget(face)
                    check.add_widget(checkbox)
                    item.add_widget(check)
                    b = str(branch.title)+'email'
                    self.ids[b].add_widget(item)
                except:pass

            mdlst_employee.add_widget(panel)
        self.ids['box_email'].add_widget(scroll)

    def no_employee_label(self):
        for path in os.listdir(f'slip/{SHOP_NAME}'):
            try:
                print(f'slip/{SHOP_NAME}/{path}')
                os.rmdir(f'slip/{SHOP_NAME}/{path}')
            except:pass
        label = MDLabel(
                text= 'No employee deploy',
                halign= 'center',
                pos_hint={"center_y":.5},
                theme_text_color='Hint')
        self.ids['no_employee_label_email'] = label
        self.ids['box_email'].add_widget(label)

class Employee(MDScreen):
    recent_employee = 0
    total_individual_checkbox = 0
    going_to_delete = []

    def on_start(self):
        global employee_interval
        employee_interval = Clock.schedule_interval(self.update_lst,0.5)

    def update_lst(self,dt):
        if creating == True:
            return
        
        f = []
        for path in os.listdir(f'slip/{SHOP_NAME}'):
            for i in os.listdir(f'slip/{SHOP_NAME}/'+path):
                f.append(i)
        
        if len(f) != self.recent_employee:
            self.ids['box_employee'].clear_widgets()
            if len(f) != 0:
                Clock.schedule_once(lambda dt: self.add_lst())
            else:
                Clock.schedule_once(lambda dt: self.no_employee_label())
                self.ids.delete_button.disabled=True
                self.ids.delete_label_count.text = '0'
            self.recent_employee = len(f)

    def deletion(self, path):
        try:
            os.remove(path)
        except Exception as e:
            print(e)

    def delete_selected(self):
        for person in self.going_to_delete:
            print(person)
            thread = threading.Thread(target=self.deletion, args=(person['path'],))
            thread.start()      
        snackbar = MDSnackbar(
                MDLabel(
                text="Deleted !",
                font_name = 'sarabun',
                text_color=self.theme_cls.opposite_bg_normal,
            ),
            pos_hint={"center_x": 0.5},
            size_hint_x=0.5,
            md_bg_color="#ffffff",
            duration = 3
            )
        snackbar.open()

    def individual_selected(self,button:MDCheckbox,pos):
        data = button.parent.parent.parent
        payload={
            'path':data.path,
            'name':data.name,
            'email':data.email,
            'branch':data.branch,
            'file_name':data.file_name
        }
        if button.active:
            self.going_to_delete.append(payload)
        else:
            try:
                self.going_to_delete.remove(payload)
            except:pass
        true_state = len(self.going_to_delete)

        self.ids.delete_label_count.text = str(true_state)

        if true_state > 0:
            self.ids.delete_button.disabled=False
        else:
            self.ids.delete_button.disabled=True

        if true_state == self.total_individual_checkbox:
            self.ids.all_select_employee.active = True
        else:
            self.ids.all_select_employee.active = False


    def all_selected(self,button:MDRaisedButton):
        state = button.active
        for ids in self.ids: 
            if 'checkbox_employee_' in ids:
                self.ids[ids].active = state
                
    
    def add_lst(self):
        self.going_to_delete=[]
        self.ids.delete_label_count.text = str(len(self.going_to_delete))
        fram1 = MDBoxLayout(
            orientation="horizontal",
            size_hint= (1,0.08)
        )
        self.ids['fram1_employee'] = fram1
        fram1.add_widget(
            MDLabel(
                text= "All",
                size_hint= (0.3,1),
                pos_hint={'center_y':0.75},
                halign= 'right',
                font_style= "sarabun",
                font_size= "25dp"
            )
        )
        all_checkbox = MDCheckbox(
            size_hint= (0.3,1),
            pos_hint={'center_y':0.65},
            halign= 'left',
        ) 
        all_checkbox.bind(on_press=self.all_selected)
        self.ids["all_select_employee"] = all_checkbox
        fram1.add_widget(all_checkbox)
        self.ids['box_employee'].add_widget(fram1)
        scroll = MDScrollView()
        self.ids['scroll_employee'] = scroll
        mdlst_employee = MDList()
        scroll.add_widget(mdlst_employee)

        for branch in os.listdir(f'slip/{SHOP_NAME}'):
            content = MDBoxLayout()
            content.adaptive_height = True
            content.orientation = 'vertical'
            self.ids[f"{branch.title}employee"] = content 
            panel = MDExpansionPanel(
                icon="office-building-outline",
                content=content,
                pos_hint={'center_y':1},
                panel_cls=MDExpansionPanelTwoLine(
                    text=f"[size=25]{branch}[/size]",
                    secondary_text=f"{len(os.listdir(f'slip/{SHOP_NAME}/'+branch))}คน",
                    font_style= "sarabunBold",
                    secondary_font_style= "sarabunBold"
                )
            )
            for num,filename in enumerate(os.listdir(f'slip/{SHOP_NAME}/'+branch),1):
                    try:
                        name,email,checker,ofmonth,createat = filename[:-4].split(',')
                        item = CustomOneLineAvatarIconListItem(
                            text=f"[size=20]{num}.{name}   |   {email}   |   Salary of [b]{ofmonth}[/b]   [color=9C9B9B]Create at {datetime.strptime(createat,'%d%m%y%H%M%S').strftime('%d/%m/%y %H:%M')}[/color][/size]",
                            font_style = "sarabun",
                            path= os.path.join('slip',SHOP_NAME,branch,filename),
                            email=filename.split(',')[1],
                            name=filename.split(',')[0],
                            branch=branch,
                            file_name=filename
                        )
                        face = IconLeftWidget(
                            icon='account'
                        )
                        check = list_container()
                        checkbox = MDCheckbox()
                        checkbox.bind(active=self.individual_selected)
                        self.total_individual_checkbox += 1
                        self.ids[f"checkbox_employee_{datetime.now().strftime('%f')}"] = checkbox
                        checkbox.color_active = self.theme_cls.accent_light
                        item.add_widget(face)
                        check.add_widget(checkbox)
                        item.add_widget(check)
                        b = str(branch.title)+'employee'
                        self.ids[b].add_widget(item)
                    except:pass
            mdlst_employee.add_widget(panel)
        self.ids['box_employee'].add_widget(scroll)

    def no_employee_label(self):
        for path in os.listdir(f'slip/{SHOP_NAME}'):
            try:
                os.rmdir(f'slip/{SHOP_NAME}/{path}')
            except:pass
        label = MDLabel(
                text= 'No employee deploy',
                halign= 'center',
                pos_hint={"center_y":.5},
                theme_text_color='Hint')
        self.ids['no_employee_label'] = label
        self.ids['box_employee'].add_widget(label)

class Setting(MDScreen):
    def on_start(self):
        pass

class SliperApp(MDApp):
    employee = []
    def build(self):
        self.title = TITLE
        self.icon = f'data\icon\{SHOP_NAME}.png'
        LabelBase.register(name='sarabun',
            fn_regular=r'data\font\THSarabunNew.ttf',
        )
        theme_font_styles.append('sarabun')
        self.theme_cls.font_styles["sarabun"] = [
            "sarabun",
            20,
            False,
            0.15,
        ]
        LabelBase.register(name='sarabunBold',
            fn_regular=r'data\font\THSarabunNew Bold.ttf'
        )
        theme_font_styles.append('sarabunBold')
        self.theme_cls.font_styles["sarabunBold"] = [
            "sarabunBold",
            20,
            False,
            0.15,
        ]
        self.theme_cls.primary_palette = "Brown"
        self.theme_cls.primary_hue = "400"
        self.theme_cls.theme_style = "Light"
        screen = Builder.load_file('App.kv')
        return screen
    
    def on_start(self):
       if not os.path.exists(f'slip'):
            os.mkdir(f'slip')

       if not os.path.exists(f'slip/{SHOP_NAME}'):
            os.mkdir(f'slip/{SHOP_NAME}')
       sc_lst = self.root.ids.mana.screen_names
       for screen in sc_lst:
           self.root.ids.mana.get_screen(screen).on_start()


SliperApp().run()
