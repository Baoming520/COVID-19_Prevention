#!/usr/bin/python
from utils.logit import Logit, EmailLogit, InfoLogit
import datetime
import json
import os
import shutil

logname = 'fileops.log'
logfile = os.path.join('./logs', logname)

# Read config file
with open('./config.json', 'r', encoding='utf-8') as f:
    config = json.load(f)

def get_file_with_latest_version(dir, fn_kw=''):
    l_file = ''
    l_time = 0
    for f in os.listdir(dir):
        f_modified_time = os.path.getmtime(os.path.join(dir, f))
        if fn_kw and fn_kw in f or not fn_kw:
            if f_modified_time > l_time:
                l_file = f
                l_time = f_modified_time

    return l_file, l_time

@Logit(logFile=logfile)
def move_files(s_dir, t_dir, fn_kw=''):
    for f in os.listdir(s_dir):
        fpath = os.path.join(t_dir, f)
        # If the file exists in the target folder, remove it. 
        if os.path.exists(fpath):
            os.remove(fpath)
        if fn_kw and fn_kw in f or not fn_kw:
            shutil.move(os.path.join(s_dir, f), t_dir)

@InfoLogit(logFile=logfile)
@EmailLogit([ config['mail_to'] ], 'Copy download data file failure')
def copy_file(s_dir, t_dir, fn_kw):
    l_f, _ = get_file_with_latest_version(s_dir, fn_kw)
    f_updated_dt = datetime.datetime.fromtimestamp(_)
    curr = datetime.datetime.now()
    delta_t = (curr - f_updated_dt).seconds
    if delta_t > 180:
        raise Exception('文件版本异常，该文件创建时间已逾180秒，请重新下载')

    print('文件是最新的，创建时间在{}秒之内。'.format(delta_t))

    # Hard code to verify the extension name of the file
    if l_f.endswith('.xlsx'):
        shutil.copyfile(os.path.join(s_dir, l_f),
                        os.path.join(t_dir, fn_kw + '.xlsx'))
        return True
    else:
        return False