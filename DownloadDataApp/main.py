#!/usr/bin/python
from email.mime.text import MIMEText
from email.header import Header
from selenium import webdriver
from selenium.webdriver.common.keys import Keys

import datetime
import json
import os
import re
import requests
import schedule
import shutil
import smtplib
import time
import uuid
import zipfile

# Read config file
with open('./config.json', 'r', encoding='utf-8') as f:
  config = json.load(f)

def get_latest_version(t_dir):
  l_file = ''
  l_time = 0
  for f in os.listdir(t_dir):
    f_modified_time = os.path.getmtime(os.path.join(t_dir, f))
    if f_modified_time > l_time:
      l_file = f
      l_time = f_modified_time
  return l_file, l_time

def get_latest_version_v2(t_dir, fn_kw):
  l_file = ''
  l_time = 0
  for f in os.listdir(t_dir):
    if fn_kw in f:
      f_modified_time = os.path.getmtime(os.path.join(t_dir, f))
      if f_modified_time > l_time:
        l_file = f
        l_time = f_modified_time
  return l_file, l_time

def move_files(s_dir, t_dir, fn_kw):
  for f in os.listdir(s_dir):
    if fn_kw in f:
      fpath = os.path.join(t_dir, f)
      if os.path.exists(fpath):
        os.remove(fpath)
      shutil.move(os.path.join(s_dir, f), t_dir)

def copy_file(s_dir, t_dir, fn_kw):
  l_f, _ = get_latest_version_v2(s_dir, fn_kw)
  f_updated_dt = datetime.datetime.fromtimestamp(_)
  curr = datetime.datetime.now()
  delta_t = (curr - f_updated_dt).seconds
  if delta_t > 180:
    failure_message = '文件版本异常，该文件创建时间已逾180秒，请重新下载'
    log(config['log_folder'], failure_message + '\n')
    send_mail(failure_message) # TODO: 被发送的机率很小
    return False
  succ_message = '文件是最新的，创建时间在{}秒之内。'.format(delta_t)
  log(config['log_folder'], succ_message + '\n')
  # send_mail(succ_message)

  # Hard code to verify the extension name of the file
  if l_f.endswith("xlsx"):
    shutil.copyfile(os.path.join(s_dir, l_f),
                    os.path.join(t_dir, fn_kw + '.xlsx'))
    return True
  else:
    return False

def log(log_dir, msg):
  with open(os.path.join(log_dir, 'download.log'), 'a+') as f:
    f.write('{}:{}'.format(
        datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S'), msg))

def send_mail(message):
  m_from = config['mail_from']
  m_to_arr = [config['mail_to']]
  html = '<p>{}</p>'.format(message)
  m_msg = MIMEText(html, 'html', 'utf-8')
  m_msg['From'] = Header(config['mail_from_displayname'], 'utf-8')
  m_msg['To'] = Header('Download Report', 'utf-8')
  m_msg['Subject'] = Header('Download Report', 'utf-8')
  try:
    o_smtp = smtplib.SMTP()
    o_smtp.connect(config['mail_host'], config['mail_port'])
    o_smtp.login(m_from, config['mail_pass'])
    o_smtp.sendmail(m_from, m_to_arr, m_msg.as_string())
  except smtplib.SMTPException:
    log(config['log_folder'], '邮件发送失败！')

def get_chrome_version():
  ret = None
  local = config['chrome_local']
  for parent, dirnames, filename in os.walk(local):
    for dirname in dirnames:
      m = re.match('[0-9]+(.[0-9]+)*', dirname)
      if m != None:
        ret = m.string
        break
    break

  return ret

def check_version():
  validation = False
  # Get the version of the local Chrome browser
  chrome_version = get_chrome_version()
  print('Local Chrome\'s Version: ' + chrome_version)
  try:
    # Try to load the chrome driver
    web_driver_path = os.path.join(
        config['web_driver_dir'], config['web_driver_name'])
    with webdriver.Chrome(web_driver_path) as driver:
      # Try to get the version of chrome driver
      driver_version = driver.capabilities['chrome']['chromedriverVersion']
      m = re.match('[0-9]+(.[0-9]+)*', driver_version)
    if m != None:
      driver_version = m.group()
      print('Chrome Driver\'s Version: ' + driver_version)
    if chrome_version == driver_version:
      validation = True
  except Exception as ex:
    print(ex)
    log(config['log_folder'], str(ex) + '\n')
    validation = False

  return (validation, chrome_version)

def unzip_file(zip_src, zip_dst):
  if zipfile.is_zipfile(zip_src):
    fz = zipfile.ZipFile(zip_src, 'r')
    for f in fz.namelist():
      fz.extract(f, zip_dst)

    return True    
  else:
    return False

def upgrade_driver(chrome_version):
  try:
    new_dir = os.path.join(config['web_driver_dir'], chrome_version)
    if not os.path.exists(new_dir):
      os.makedirs(new_dir)
      print('[created] ' + new_dir)
    online_driver_path = config['chrome_driver_mirrors'] + '/' + chrome_version + '/' + config['chrome_driver_zip']
    r = requests.get(online_driver_path)
    local_driver_path = os.path.join(config['web_driver_dir'], chrome_version, config['chrome_driver_zip'])
    with open(local_driver_path, "wb") as f:
      f.write(r.content)

    # Unzip chrome driver
    unzip_file(local_driver_path, config['web_driver_dir'])
    print('[success] Finished to upgrade chrome driver!')
    log(config['log_folder'], 'Finished to upgrade chrome driver!\n')

    return True
  except Exception as ex:
    print('[failed] Failed to upgrade chrome driver')
    log(config['log_folder'], str(ex) + '\n')

    return False

def job():
  # Check the consistency of the two versions ("Chrome Browser" and "Chrome driver")
  va, ver = check_version()
  succ = True # Means there is no need to upgrade the chrome driver
  if not va:
    print('更新数据中，请稍后……')
    succ = upgrade_driver(ver)
  
  if not succ:
    print('Chrome Driver更新失败！')
    return
  
  # Load Chrome driver
  web_driver_path = os.path.join(config['web_driver_dir'], config['web_driver_name'])
  wd = webdriver.Chrome(web_driver_path)
  try:
    wd.get(config['website'])

    # Click the "login and edit" button to login
    wd.implicitly_wait(config['loading_timeout']) # Set the timeout for loading page
    login_btn = wd.find_element_by_xpath('//div[@class="component-login-modal-pop"]/div[@class="modal-wrap s"]/div[@class="component-text-btn system"]')
    login_btn.click() # Go to the login options page

    # Wait for loading the login options page
    wd.implicitly_wait(config['loading_timeout']) # Set the timeout for loading page
    print('INFO: 登陆选项页面加载成功')

    # Accept the user login protocol on login options page
    wd.find_element_by_id('loginProtocal').click()

    # Choose the login option with an existing account
    wd.find_element_by_xpath('//span[@class="js_toProtocolDialog tpItem tpItem_account"]').click() # Go to account login page

    wd.implicitly_wait(config['loading_timeout']) # Set the timeout for loading page
    print('INFO: 账号登陆页面加载成功')

    # Input email or phone number
    wd.find_element_by_id('email').send_keys(config['username'])
    # Input password
    wd.find_element_by_id('password').send_keys(config['password'])
    # Execute the verification
    wd.find_element_by_id('rectMask').click()
    time.sleep(config['normal_timeout']) # Wait for verification for a few seconds
    # Remove the keep online mode
    wd.find_element_by_id('keepOnline').click()
    # Click the login button
    wd.find_element_by_id('login').click() # Go to the personal info page

    # Wait for loading personal info page
    wd.implicitly_wait(config['loading_timeout']) # Set the timeout for loading page
    print('INFO: 返回数据页面')
    # menu_btn = wd.find_element_by_id('header-more-btn') # There is no 'id' property in the menu-btn (div).
    menu_btn = wd.find_element_by_xpath('//div[@class="component-icon-btn header-more-btn"]')
    time.sleep(config['normal_timeout']) # Wait for loading the actions for the menu button
    menu_btn.click()

    wd.implicitly_wait(config['loading_timeout']) # Set the timeout for loading page
    # menu_items = wd.find_elements_by_xpath('//div[@class="component-menu-item only-text"]')
    menu_items = wd.find_elements_by_xpath('//div[@class="menu-text"]')
    for it in menu_items:
      if it.text == '下载':
        it.click()
        time.sleep(config['download_timeout']) # Wait for downloading the excel file
        break
  except Exception as ex:
    print(ex)
    log(config['log_folder'], str(ex) + '\n')
    send_mail('Unknown Issue: ' + str(ex) + '\n')
    return
  finally:
    wd.close()

  # move_files(config['download_folder'], config['move_folder'], config['download_filename_kw'])
  # l_f, l_t = get_latest_version(config['move_folder'])
  # for rm_f in os.listdir(config['move_folder']):
  #   if rm_f != l_f:
  #     os.remove(os.path.join(config['move_folder'], rm_f))

  # Copy the target file from download folder to destination
  res = copy_file(config['download_folder'], config['move_folder'], config['download_filename_kw'])
  if res:
    l_f, l_t = get_latest_version(config['move_folder'])
    log(config['log_folder'], '下载成功！\n')
    print('下载数据完成，本地数已是最新')
    print('数据更新时间为{}'.format(datetime.datetime.fromtimestamp(l_t)))
    print('Done!')
  else:
    print('下载数据失败！')
    log(config['log_folder'], '下载失败！\n')
    send_mail('下载数据失败！')
    

def main(argv=None):
  job()
  # schedule.every(1).minutes.do(job)
  # schedule.every().hour.do(job)
  # while True:
  #   schedule.run_pending()
  #   time.sleep(1)

if __name__ == '__main__':
  main()
  
