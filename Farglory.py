## Import Package
import pandas as pd
import numpy as np
pd.options.display.max_rows = 10

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import NoSuchElementException 
from selenium.webdriver import ActionChains
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By

import requests
from bs4 import BeautifulSoup
from bs4 import NavigableString, Tag

import time
import datetime
today = datetime.date.today() #20220731
last_month_end = today.replace(day=1) - datetime.timedelta(days=1)

yyyymm_today = int(today.strftime('%Y%m'))
yyyymm_last  = int(last_month_end.strftime('%Y%m'))

import os

path = r'C:/Users/003094/Desktop/Python/同業宣告利率'
os.chdir(path)

company_url = pd.read_excel('宣告網址.xlsx')

print(yyyymm_today)
print(yyyymm_last)


#############################################################################




# 取得網址

company = '遠雄'
url = company_url.loc[company_url['壽險公司'] == company, '宣告網址'].values[0]
print(f'{company}宣告利率爬蟲開始')




chrome_browser = webdriver.Chrome()

chrome_browser.get(url)
time.sleep(3)  

# 關掉通知
button = chrome_browser.find_element_by_xpath('/html/body/div[2]/div[2]/div[3]/div[1]/div/div/div/a')
ActionChains(chrome_browser).move_to_element(button).click(button).perform()
time.sleep(3)  

### 這個月

# 選年
year = str(int(yyyymm_today / 100) - 1911) #民國年
select = Select(chrome_browser.find_element_by_id('selYear'))
select.select_by_visible_text(year)
time.sleep(2)

# 選月      
month = str(int(yyyymm_today % 100))
select = Select(chrome_browser.find_element_by_id('selMonth'))
select.select_by_visible_text(month)
time.sleep(2)

# 爬
page = chrome_browser.page_source
page_soup = BeautifulSoup(page,'html.parser')

body = page_soup.findAll('tbody')
containers1 = body[1].findAll('tr')

df_today = pd.DataFrame(columns = ['商品名稱', f'{yyyymm_today}宣告利率'])

for i in range(len(containers1)-1):
    df_today.loc[i,:] = [containers1[i].findAll('td')[0].text,
                         containers1[i].findAll('td')[1].text]
    
    
    
    
## 上個月

# 選年
year = str(int(yyyymm_last / 100) - 1911)
select = Select(chrome_browser.find_element_by_id('selYear'))
select.select_by_visible_text(year)
time.sleep(2)

# 選月      
month = str(int(yyyymm_last % 100))
select = Select(chrome_browser.find_element_by_id('selMonth'))
select.select_by_visible_text(month)
time.sleep(2)

# 爬
page = chrome_browser.page_source
page_soup = BeautifulSoup(page,'html.parser')

body = page_soup.findAll('tbody')
containers1 = body[1].findAll('tr')

df_last = pd.DataFrame(columns = ['商品名稱', f'{yyyymm_last}宣告利率'])

for i in range(len(containers1)-1):
    df_last.loc[i,:] = [containers1[i].findAll('td')[0].text,
                   containers1[i].findAll('td')[1].text]
    
chrome_browser.close()


df = df_today.merge(df_last, how='left', on=['商品名稱'])
df
