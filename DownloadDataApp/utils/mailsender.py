#!/usr/bin/python
from email.mime.text import MIMEText
from email.header import Header

import smtplib

class MailSender():
    def __init__(self, mailHost, mailPort, mailFrom, mailFromDisplayName, mailPassword):
        self.mailHost = mailHost
        self.mailPort = mailPort
        self.mailFrom = mailFrom
        self.mailFromDisplayName = mailFromDisplayName
        self.mailPassword = mailPassword
    
    def send(self, mailToList, subject, content):
        html = '<p>{}</p>'.format(content)
        mailMsg = MIMEText(html, 'html', 'utf-8')
        mailMsg['From'] = Header(self.mailFromDisplayName, 'utf-8')
        # mailMsg['To'] = Header('', 'utf-8')
        mailMsg['Subject'] = Header(subject, 'utf-8')
        try:
            smtp = smtplib.SMTP()
            smtp.connect(self.mailHost, self.mailPort)
            smtp.login(self.mailFrom, self.mailPassword)
            smtp.sendmail(self.mailFrom, mailToList, mailMsg.as_string())
        except smtplib.SMTPException as smtpEx:
            raise smtpEx
