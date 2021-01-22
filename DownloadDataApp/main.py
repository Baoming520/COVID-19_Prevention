#!/usr/bin/python
from helper.filehelper import copy_file, get_file_with_latest_version
from helper.mailhelper import send_mail
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait as WDWait

from utils.browserdriver import \
    check_version_consistency, \
    download_chrome_driver, \
    get_chrome_driver_version_list, \
    choose_matching_version, \
    upgrade_chrome_driver
from utils.logit import Logit, InfoLogit, SuccLogit

import datetime
import json
import os
import schedule
import time

# Read config file
with open('./config.json', 'r', encoding='utf-8') as f:
    config = json.load(f)

def log(log_dir, msg):
    with open(os.path.join(log_dir, 'download.log'), 'a+') as f:
        f.write('{}:{}'.format(
            datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S'), msg))

def job():
    # Check the consistency of the two versions ("Chrome Browser" and "Chrome driver")
    passed, ver = check_version_consistency(config['chrome_local'], os.path.join(
        config['web_driver_dir'], config['web_driver_name']))
    if not passed:
        res, upgrade_msg = upgrade_chrome_driver(config['chrome_driver_mirrors'], ver)
        print(upgrade_msg)
    
        if not res:
            print('更新驱动失败，终止运行！')
            return

    # Load Chrome driver
    web_driver_path = os.path.join(config['web_driver_dir'], config['web_driver_name'])
    if config['chrome_headless_mode']:
        # Config with headless mode
        chrome_options = Options()
        chrome_options.add_argument('--headless')
        wd = webdriver.Chrome(web_driver_path, chrome_options=chrome_options)
    else:
        wd = webdriver.Chrome(web_driver_path)

    try:
        wd.get(config['website'])

        # Confirmation: Invite to cooperate
        wd.find_element_by_xpath(
            '//div[@class="preview-error-container coopInvite"]/a/button[@class="el-button btn btn-allow el-button--primary"]').click()

        # Wait for loading the login options page
        # Set the timeout for loading page
        wd.implicitly_wait(config['loading_timeout'])
        print('INFO: 登陆选项页面加载成功')
        
        # Agree the protocol and policy
        agree_checkbox = wd.find_element_by_id('loginProtocal')
        agree_checkbox.click()

        # Click the "login and edit" button to login
        # Set the timeout for loading page
        wd.implicitly_wait(config['loading_timeout'])

        # Choose the login option with an existing account
        # Go to account login page
        # login_btn = wd.find_element_by_xpath(
        #     '//div[@class="component-login-modal-pop"]/div[@class="modal-wrap s"]/div[@class="component-text-btn system"]')
        login_btn = wd.find_element_by_xpath(
            '//div[@id="mainWrap"]/div[@class="keepOnline_wrap"]/a[@data-to="accountWrap"]')
        login_btn.click()  # Go to the login options page

        # Set the timeout for loading page
        wd.implicitly_wait(config['loading_timeout'])
        print('INFO: 账号登陆页面加载成功')

        # Input email or phone number
        wd.find_element_by_id('email').send_keys(config['username'])
        # Input password
        wd.find_element_by_id('password').send_keys(config['password'])
        # Execute the verification
        wd.find_element_by_id('rectMask').click()
        # Wait for verification for a few seconds
        ver_text = wd.find_element_by_id('SM_TXT_1').text
        while ver_text != '验证成功':
            time.sleep(config['normal_timeout'])
            ver_text = wd.find_element_by_id('SM_TXT_1').text
        # Remove the keep online mode
        wd.find_element_by_id('keepOnline').click()
        # Click the login button
        loginBtn = wd.find_element_by_id('login')
        # loginBtn.click()  # Go to the personal info page
        wd.execute_script('arguments[0].click();', loginBtn)
        print('INFO: 返回数据页面')

        # Wait for the data page is loaded completely
        print('[START]: Waiting for the specified element')
        WDWait(wd, 20).until(EC.presence_of_all_elements_located((By.ID, 'root'))) # Waiting for the menu button
        WDWait(wd, 20).until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, '.component-icon-btn.header-more-btn'))) # The class name has more than one values
        print('[END]: Found the element')

        time.sleep(config['normal_timeout'] + 5)
        
        menu_btn = wd.find_element_by_xpath(
            '//div[@class="component-icon-btn header-more-btn"]')
        menu_btn.click()
        
        # Set the timeout for loading page
        wd.implicitly_wait(config['loading_timeout'])
        # menu_items = wd.find_elements_by_xpath('//div[@class="component-menu-item only-text"]')
        menu_items = wd.find_elements_by_xpath('//div[@class="menu-text"]')
        for it in menu_items:
            if it.text == '下载':
                it.click()
                # Wait for downloading the excel file
                time.sleep(config['download_timeout'])
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
    res = copy_file(config['download_folder'],
                    config['move_folder'], config['download_filename_kw'])
    if res:
        l_f, l_t = get_file_with_latest_version(config['move_folder'])
        log(config['log_folder'], '下载成功！\n')
        print('下载数据完成，本地数已是最新')
        print('数据更新时间为{}'.format(datetime.datetime.fromtimestamp(l_t)))
        print('Done!')
    else:
        print('下载数据失败！')
        log(config['log_folder'], '下载失败！\n')
        send_mail('下载数据失败！')

logname = 'dev_test.log'
logfile = os.path.join('./logs', logname)

# @InfoLogit(logFile=logfile) # second
# @SuccLogit(logFile=logfile) # first
# @Logit(logFile=logfile)
# def dev_test(x, y):
#     return x + y

def main(argv=None):
    job()

    # schedule.every(1).minutes.do(job)
    # schedule.every().hour.do(job)
    # while True:
    #   schedule.run_pending()
    #   time.sleep(1)
    pass


if __name__ == '__main__':
    main()
