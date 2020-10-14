#!/usr/bin/python
from bs4 import BeautifulSoup
from selenium import webdriver
from utils.logit import Logit
import json
import os
import re
import requests
import zipfile

# Use the following link to download chrome driver for outside:
# https://chromedriver.storage.googleapis.com/index.html
# Use the following link to download chrome driver for China Mainland:
# https://npm.taobao.org/mirrors/chromedriver
# You should change them in the config.json file.

logname = 'browserdriver.log'
logfile = os.path.join('./logs', logname)
driverdir = './drives'
drivername = 'chromedriver.exe'
driverzippackagename = 'chromedriver_win32.zip'


@Logit(logfile)
def upgrade_chrome_driver(driver_mirrors, browser_version):
    msg = ''
    verlist = get_chrome_driver_version_list(driver_mirrors)
    ver = choose_matching_version(verlist, browser_version)
    if download_chrome_driver(driver_mirrors, ver):
        if unzip_chrome_driver_package(os.path.join(driverdir, ver, driverzippackagename), driverdir):
            msg = 'Chrome驱动升级成功！'
            return True, msg
        else:
            msg = 'Chrome驱动压缩包解压失败！'
            return False, msg
    else:
        msg = 'Chrome驱动下载失败！'

    return False, msg


@Logit(logfile)
def check_version_consistency(installation_path, driver_path):
    validation = False
    # Get the version of the local Chrome browser
    browser_version = get_chrome_browser_version(installation_path)
    if not browser_version:
        msg = '未找到Chrome浏览器\n'
        print(msg)
        raise Exception(msg)

    print('Chrome浏览器的版本是：{}\n'.format(browser_version))
    driver_version = get_chrome_driver_version(driver_path)
    if not driver_version:
        msg = '未能获取Chrome驱动的版本\n'
        print(msg)
        validation = False  # Add it only for easy reading

    print('Chrome驱动的版本是：{}\n'.format(driver_version))
    if browser_version and driver_version and browser_version[:browser_version.rindex('.')] == driver_version[:driver_version.rindex('.')]:
        validation = True

    return (validation, browser_version)


@Logit(logfile)
def get_chrome_driver_version_list(driver_mirrors):
    version_list = []
    r = requests.get(driver_mirrors)
    if r and r.status_code == 200:
        soup = BeautifulSoup(r.text, 'html.parser', from_encoding='utf-8')
        links = soup.find_all('a')
        for link in links:
            m = re.match('^[0-9]+(.[0-9]+)*/$', link.next)
            if m != None:
                version_list.append(m.string.strip('/'))
        return version_list
    else:
        raise Exception(
            'Cannot get the version list of chrome driver in the website "{}"'.format(driver_mirrors))

# The function "choose_matching_version" should not throw any exceptions.
# So there is no need to add any decorators!


def choose_matching_version(version_list, expected_version):
    ret = ''
    maxV = 0
    for version in version_list:
        v = version[:version.rindex('.')]
        if expected_version.startswith(v):
            av_arr = [int(x) for x in version.split('.')]
            ev_arr = [int(x) for x in expected_version.split('.')]
            if len(av_arr) != len(ev_arr):
                ret = '.'.join(av_arr)
            else:
                lIdx = len(av_arr) - 1
                if av_arr[lIdx] == ev_arr[lIdx]:
                    ret = '.'.join(av_arr)
                    return ret
                else:
                    if maxV < av_arr[lIdx]:
                        maxV = av_arr[lIdx]
                        ret = version
    return ret


@Logit(logfile)
def download_chrome_driver(driver_mirrors, version):
    newdir = os.path.join(driverdir, version)
    if not os.path.exists(newdir):
        os.makedirs(newdir)
        print('新建路径："{}"'.format(newdir))

    # Download the latest version of chrome driver.
    destUrl = '{}/{}/{}'.format(driver_mirrors, version, driverzippackagename)
    r = requests.get(destUrl)
    if r and r.status_code == 200:
        local_driver_path = os.path.join(
            driverdir, version, driverzippackagename)
        with open(local_driver_path, "wb") as f:
            f.write(r.content)
        return True
    else:
        return False


@Logit(logfile)
def get_chrome_browser_version(installation_path):
    ret = None
    for parent, dirnames, filename in os.walk(installation_path):
        for dirname in dirnames:
            m = re.match('[0-9]+(.[0-9]+)*', dirname)
            if m != None:
                ret = m.string
                break
        break

    return ret


@Logit(logfile)
def get_chrome_driver_version(driver_path):
    # Try to load the chrome driver
    web_driver_path = os.path.join(driver_path)
    with webdriver.Chrome(web_driver_path) as driver:
        # Try to get the version of chrome driver
        driver_version = driver.capabilities['chrome']['chromedriverVersion']
        m = re.match('[0-9]+(.[0-9]+)*', driver_version)
        if m != None:
            driver_version = m.group()
        if not driver_version:
            raise Exception(
                'Cannnot get the chrome driver\'s version, for details: \n')
    return driver_version


@Logit(logfile)
def unzip_chrome_driver_package(srcfile, dest):
    if zipfile.is_zipfile(srcfile):
        fz = zipfile.ZipFile(srcfile, 'r')
        for f in fz.namelist():
            fz.extract(f, dest)
        return True
    else:
        return False
