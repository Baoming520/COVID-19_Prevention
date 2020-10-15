from functools import wraps
from helper.mailhelper import send_mail
import datetime
import os

logname = 'default.log'
logfile = os.path.join('./logs', logname)

class Logit(object):
    def __init__(self, logFile=logfile):
        self.logFile = logFile
    
    def __call__(self, func):
        @wraps(func)
        def decorated(*args, **kwargs):
            try:
                return func(*args, **kwargs)
            except Exception as ex:
                curr = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                err_msg = '[{}] The function "{}" has been called. Exception with following issues: \n'.format(curr, func.__name__)
                err_msg += str(ex) + '\n'
                print(err_msg)
                mode = 'a+'
                if not os.path.exists(os.path.dirname(self.logFile)):
                    os.makedirs(os.path.dirname(self.logFile)) 
                if not os.path.exists(self.logFile):
                    mode = 'w+'
                with open(self.logFile, mode) as fi:
                    fi.write(err_msg + '\n')
                    self.notify(err_msg)
                
        return decorated
    def notify(self, msg):
        # Do nothing here.
        pass

class EmailLogit(Logit):
    def __init__(self, mail_to_list, subject, *args, **kwargs):
        self.mail_to_list = mail_to_list
        self.subject = subject
        super(EmailLogit, self).__init__(*args, **kwargs)
    
    def notify(self, msg):
        # Send the error message to the specified people with mails.
        send_mail(msg, self.subject, self.mail_to_list)

class InfoLogit(object):
    def __init__(self, logFile=logfile):
        self.logFile = logFile
    
    def __call__(self, func):
        @wraps(func)
        def decorated(*args, **kwargs):
            try:
                res = func(*args, **kwargs)
                self.notify(res, func.__name__)

                return res
            except Exception as ex:
                raise ex
                
        return decorated

    def notify(self, result, funcname):
        if result is not None:
            curr = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            info_msg = '[INFO:{}] Execution result of the function "{}" is {}.\n'.format(curr, funcname, result)
            print(info_msg)
            mode = 'a+'
            if not os.path.exists(os.path.dirname(self.logFile)):
                os.makedirs(os.path.dirname(self.logFile)) 
            if not os.path.exists(self.logFile):
                mode = 'w+'
            with open(self.logFile, mode) as fi:
                fi.write(info_msg + '\n')

class SuccLogit(object):
    def __init__(self, logFile=logfile):
        self.logFile = logFile
    
    def __call__(self, func):
        @wraps(func)
        def decorated(*args, **kwargs):
            try:
                func(*args, **kwargs)
                self.notify(func.__name__)
            except Exception as ex:
                raise ex
                
        return decorated

    def notify(self, funcname):
        curr = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        succ_msg = '[SUCC:{}] The function "{}" has been executed successfully.\n'.format(curr, funcname)
        print(succ_msg)
        mode = 'a+'
        if not os.path.exists(os.path.dirname(self.logFile)):
            os.makedirs(os.path.dirname(self.logFile)) 
        if not os.path.exists(self.logFile):
            mode = 'w+'
        with open(self.logFile, mode) as fi:
            fi.write(succ_msg + '\n')