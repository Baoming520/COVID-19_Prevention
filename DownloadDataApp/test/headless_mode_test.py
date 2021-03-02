from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait as WDWait

# Switch
headless_mode = True # set it to True, if uses the headless mode, otherwise False

# Constants
timeout = 20
waiting_time = 5

if not headless_mode:
    wd = webdriver.Chrome()
else:
    # Start the chrome driver with headless mode
    chrome_options = Options()
    chrome_options.add_argument('--headless')
    wd = webdriver.Chrome(chrome_options=chrome_options)

wd.get("https://www.baidu.com/")
print(wd.title)

inputBox = wd.find_element_by_id('kw')
inputBox.send_keys('ABNF')
searchBtn = wd.find_element_by_id('su')
searchBtn.click()
# wd.implicitly_wait(5)
WDWait(wd, timeout).until(EC.presence_of_all_elements_located((By.ID, '1')))  # there is no redirect action, so please use the explicit waiting method
print(wd.title)
results = wd.find_elements_by_xpath('//div[@class="result c-container new-pmd"]')
curr_win_handle = wd.current_window_handle
if len(results) > 0:
    index = -1
    for res in results:
        index += 1
        if index == 0:
            link = res.find_element_by_xpath('./h3/a')
            link.click()
            wd.implicitly_wait(waiting_time)
            break

win_handles = wd.window_handles
for win_handle in win_handles:
    if curr_win_handle != win_handle:
        wd.switch_to_window(win_handle)
        break  

print(wd.title)
wd.close()