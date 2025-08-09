from dotenv import load_dotenv
import json
import sys
import os

load_dotenv('.env.local')
with open("config.json","r",encoding="utf8") as config:
    config = json.load(config)

if config['shop_key'] == "HARIS":
    SENDER = os.getenv('sender_email_haris')
    PASSWORD = os.getenv('email_password_haris')
    SHOP_NAME = 'haris'
    FROM_EMAIL = "Haris premium buffet"
    TITLE = "Haris Slipper"
    LOGO_PATH = "data\image\Harislogo.jpg"
    
elif config['shop_key'] == "TUKKAE":
    SENDER = os.getenv('sender_email_tukkae')
    PASSWORD = os.getenv('email_password_tukkae')
    SHOP_NAME = 'tukkae'
    FROM_EMAIL = 'ตุ๊กแกอวกาศ "Steak"'
    TITLE = "Tukkae Slipper"
    LOGO_PATH = "data\image\Tukkea.jpg"

else:
    print('Invalid shop Id')
    sys.exit()

SHOP_KEY = config['shop_key']
EMAIL_ATTEMP = config['email_attemps']
SLIP_DETAIL = config[f'{SHOP_NAME}_slip_details']