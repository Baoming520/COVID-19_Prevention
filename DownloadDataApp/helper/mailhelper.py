#!/usr/bin/python
from utils.mailsender import MailSender

import json

# Read config file
with open('./config.json', 'r', encoding='utf-8') as f:
    config = json.load(f)

def send_mail(message, subject='Download Report', to_list=[ config['mail_to'] ]):
    ms = MailSender(config['mail_host'], config['mail_port'], config['mail_from'], config['mail_from_displayname'], config['mail_pass'])
    ms.send(to_list, subject, message)
