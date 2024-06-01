from dotenv import load_dotenv
import json
import sys
import os

load_dotenv('.env.local')
with open("config.json","r",encoding="utf8") as config:
            config = json.load(config)

EMAIL_ATTEMP = config['email_attemps']
SHOP_ID = config['shop_id']
if config['shop_id'] == "1":
    SENDER = os.getenv('sender_email_haris')
    PASSWORD = os.getenv('email_password_haris')
    SHOP_NAME = 'haris'
    TITLE = "Haris Slipper"
    
elif config['shop_id'] == "2":
    SENDER = os.getenv('sender_email_tukkae')
    PASSWORD = os.getenv('email_password_tukkae')
    SHOP_NAME = 'tukkae'
    TITLE = "Tukkae Slipper"
else:
    print('Invalid shop Id')
    sys.exit()