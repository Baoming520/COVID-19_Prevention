from functools import wraps
import datetime
import os

class Logit(object):
    def __init__(self, logFile):
        self.logFile = logFile
    
    def __call__(self, func):
        @wraps(func)
        def decorated(*args, **kwargs):
            try:
                return func(*args, **kwargs)
            except Exception as ex :
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
                
        return decorated
