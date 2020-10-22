import os
import re
import requests
import sys

def verifyURL(url):
  regex = re.compile(
    r'^(?:http|ftp)s?://' # http:// or https://
    r'(?:(?:[A-Z0-9](?:[A-Z0-9-]{0,61}[A-Z0-9])?\.)+(?:[A-Z]{2,6}\.?|[A-Z0-9-]{2,}\.?)|' #domain...
    r'localhost|' #localhost...
    r'\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3})' # ...or ip
    r'(?::\d+)?' # optional port
    r'(?:/?|[/?]\S+)$', re.IGNORECASE)

  return re.match(regex, url) is not None

def verifyFileName(filename):
  regex = re.compile(r'^[a-zA-Z0-9_.-]+$', re.IGNORECASE)

  return re.match(regex, filename) is not None

def download(url, localPath, filename):
  res = requests.get(url)
  with open(os.path.join(localPath, filename), 'wb') as f:
    f.write(res.content)

def main(argv=None):
  if argv == None or len(argv) < 3:
    print('Missing one or more required parameter(s)')
    return

  url = argv[1]
  if not verifyURL(url):
    print('The url is invalid.')
    return
  
  path = argv[2]
  if not os.path.exists(path):
    print('The local path is not exist.')
    return

  fname = argv[3]
  if not verifyFileName(fname):
    print('The file name contains invalid character(s).')
    return
  try:
    download(url, path, fname)
  except Exception as ex:
    print('Fatal error: \n' + ex)

if '__main__' == __name__:
  main(sys.argv)
  
