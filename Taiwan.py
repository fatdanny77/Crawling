## Import Package
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import NoSuchElementException 
from selenium.webdriver import ActionChains
import pandas as pd
import numpy as np
import requests
from bs4 import BeautifulSoup
from bs4 import NavigableString, Tag
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By

import time
import datetime
import os

today = datetime.date.today()
last_month_end = today.replace(day=1) - datetime.timedelta(days=1)

yyyymm_today = int(today.strftime('%Y%m'))
yyyymm_last  = int(last_month_end.strftime('%Y%m'))

company = '台灣'

print(f'{company}宣告利率爬蟲開始')
print('本月為 ' + str(yyyymm_today))
print('上月為 ' + str(yyyymm_last))


#######################################################################################



pd.options.display.max_rows = 10
pd.options.display.max_rows

company_url = pd.read_excel('C:/Users/003094/Desktop/Python/同業宣告利率/宣告網址.xlsx')
url = company_url.loc[company_url['壽險公司'] == company, '宣告網址'].values[0]
url



def get_rate(mm):
    page = chrome_browser.page_source
    page_soup = BeautifulSoup(page,'html.parser')
    body = page_soup.tbody
    containers1 = body.findAll('tr')
    rate_list = []
    for i in range(len(containers1)):
        containers2 = containers1[i].findAll('td')
        rate_list_a = [q.text.strip() for q in containers2]
        rate_list.append(rate_list_a)


    df = pd.DataFrame(rate_list, columns = ['yyyymm', '商品名稱', str(mm) +'宣告利率'])
    
    return df

  
  
  
 # 點下一頁

def get_rate_action(mm):
    page = chrome_browser.page_source
    page_soup = BeautifulSoup(page,'html.parser')

    df = get_rate(mm)
    while True:
        try:
            next_button = chrome_browser.find_element_by_xpath('/html/body/app-root/rdr-render/rdr-templates-container/rdr-dynamic-wrapper/twlife-empty-layout/twlife-layout-scroll-wrap/div/div[2]/twlife-layout-frame/div/div/twlife-layout-block/div/div/rdr-content-level-info-render/div/rdr-templates-container/rdr-dynamic-wrapper/twlife-interest-rate-template/twlife-interest-rate/div/div[2]/twlife-interest-rate-list/twlife-layout-content-block/div[2]/div[2]/rdr-pagination/div[2]')
            classAttribute = next_button.get_attribute('class')
            if 'disabled' in classAttribute:
                break
            else:
                WebDriverWait(chrome_browser, 10).until(EC.element_to_be_clickable((By.XPATH, '/html/body/app-root/rdr-render/rdr-templates-container/rdr-dynamic-wrapper/twlife-empty-layout/twlife-layout-scroll-wrap/div/div[2]/twlife-layout-frame/div/div/twlife-layout-block/div/div/rdr-content-level-info-render/div/rdr-templates-container/rdr-dynamic-wrapper/twlife-interest-rate-template/twlife-interest-rate/div/div[2]/twlife-interest-rate-list/twlife-layout-content-block/div[2]/div[2]/rdr-pagination/div[2]'))).click()
                time.sleep(2)
                df1 = get_rate(mm)
                df = df.append(df1, ignore_index=True)
                time.sleep(2)  
        except:
            break
    return df
  

def get_rate_type(yyyymm):

    # 取消彈出視窗
    button = chrome_browser.find_element_by_xpath('//*[@id="cdk-overlay-0"]/rdr-snackbar/rdr-snackbar-message/div/div[4]')
    ActionChains(chrome_browser).move_to_element(button).click(button).perform()
    time.sleep(2)
    
    # 是否為上個月
    
    if yyyymm == yyyymm_last:
        
        if str(yyyymm_today)[-2:] == '01':
            # 選上一年
            button = chrome_browser.find_element_by_xpath('/html/body/app-root/rdr-render/rdr-templates-container/rdr-dynamic-wrapper/twlife-empty-layout/twlife-layout-scroll-wrap/div/div[2]/twlife-layout-frame/div/div/twlife-layout-block/div/div/rdr-content-level-info-render/div/rdr-templates-container/rdr-dynamic-wrapper/twlife-interest-rate-template/twlife-interest-rate/div/div[1]/div/div[3]/twlife-year-month-select/div/div/div[1]/rdr-select/nx-select/div/div/a')
            ActionChains(chrome_browser).move_to_element(button).click(button).perform()
            time.sleep(2)
            
            button = chrome_browser.find_element_by_xpath('/html/body/app-root/rdr-render/rdr-templates-container/rdr-dynamic-wrapper/twlife-empty-layout/twlife-layout-scroll-wrap/div/div[2]/twlife-layout-frame/div/div/twlife-layout-block/div/div/rdr-content-level-info-render/div/rdr-templates-container/rdr-dynamic-wrapper/twlife-interest-rate-template/twlife-interest-rate/div/div[1]/div/div[3]/twlife-year-month-select/div/div/div[1]/rdr-select/nx-select/div/div[2]/div/ul/li[3]/a')
            ActionChains(chrome_browser).move_to_element(button).click(button).perform()
            time.sleep(2)            
            
        button = chrome_browser.find_element_by_xpath('/html/body/app-root/rdr-render/rdr-templates-container/rdr-dynamic-wrapper/twlife-empty-layout/twlife-layout-scroll-wrap/div/div[2]/twlife-layout-frame/div/div/twlife-layout-block/div/div/rdr-content-level-info-render/div/rdr-templates-container/rdr-dynamic-wrapper/twlife-interest-rate-template/twlife-interest-rate/div/div[1]/div/div[3]/twlife-year-month-select/div/div/div[2]/rdr-select/nx-select/div/div')
        ActionChains(chrome_browser).move_to_element(button).click(button).perform()
        time.sleep(2)
        
        # 找月份總長度
        page = chrome_browser.page_source
        page_soup = BeautifulSoup(page,'html.parser')
        ul = page_soup.findAll('ul', {'class' : 'select-options__list'})[0].findAll('li')
        length = len(ul) + 1

        # 用順序選擇月份
        month = str(yyyymm)[-2:]
        num = length - int(month)
        button = chrome_browser.find_element_by_xpath(f'/html/body/app-root/rdr-render/rdr-templates-container/rdr-dynamic-wrapper/twlife-empty-layout/twlife-layout-scroll-wrap/div/div[2]/twlife-layout-frame/div/div/twlife-layout-block/div/div/rdr-content-level-info-render/div/rdr-templates-container/rdr-dynamic-wrapper/twlife-interest-rate-template/twlife-interest-rate/div/div[1]/div/div[3]/twlife-year-month-select/div/div/div[2]/rdr-select/nx-select/div/div[2]/div/ul/li[{num}]')
        ActionChains(chrome_browser).move_to_element(button).click(button).perform()
        time.sleep(2)

    # 下拉商品類別選單
    button = chrome_browser.find_element_by_xpath('/html/body/app-root/rdr-render/rdr-templates-container/rdr-dynamic-wrapper/twlife-empty-layout/twlife-layout-scroll-wrap/div/div[2]/twlife-layout-frame/div/div/twlife-layout-block/div/div/rdr-content-level-info-render/div/rdr-templates-container/rdr-dynamic-wrapper/twlife-interest-rate-template/twlife-interest-rate/div/div[1]/div/div[1]/rdr-select/nx-select/div/div/a')
    ActionChains(chrome_browser).move_to_element(button).click(button).perform()
    time.sleep(2)

    # 取得選單長度
    page = chrome_browser.page_source
    page_soup = BeautifulSoup(page,'html.parser')
    group = page_soup.findAll('ul', {"class" : "select-options__list"})[0].findAll('li')

    # 再點一下，取消選單
    ActionChains(chrome_browser).move_to_element(button).click(button).perform()
    time.sleep(2)
        
    for i in range(2, len(group)+1):

        button = chrome_browser.find_element_by_xpath('/html/body/app-root/rdr-render/rdr-templates-container/rdr-dynamic-wrapper/twlife-empty-layout/twlife-layout-scroll-wrap/div/div[2]/twlife-layout-frame/div/div/twlife-layout-block/div/div/rdr-content-level-info-render/div/rdr-templates-container/rdr-dynamic-wrapper/twlife-interest-rate-template/twlife-interest-rate/div/div[1]/div/div[1]/rdr-select/nx-select/div/div/a')
        ActionChains(chrome_browser).move_to_element(button).click(button).perform()
        time.sleep(2)

        button1 = chrome_browser.find_element_by_xpath(f'/html/body/app-root/rdr-render/rdr-templates-container/rdr-dynamic-wrapper/twlife-empty-layout/twlife-layout-scroll-wrap/div/div[2]/twlife-layout-frame/div/div/twlife-layout-block/div/div/rdr-content-level-info-render/div/rdr-templates-container/rdr-dynamic-wrapper/twlife-interest-rate-template/twlife-interest-rate/div/div[1]/div/div[1]/rdr-select/nx-select/div/div[2]/div/ul/li[{i}]')
        group = button1.text
        chrome_browser.execute_script("arguments[0].click();", button1)
        time.sleep(2)

        if i == 2 :       
            df = get_rate_action(yyyymm)
            df['商品類別'] = group
        else:
            df1 = get_rate_action(yyyymm)
            df1['商品類別'] = group
            df = df.append(df1, ignore_index=True)
            
    return df
  
 ######################################################################


# 本月

chrome_browser = webdriver.Chrome()
chrome_browser.get(url)
time.sleep(2)

# 彈出視窗
while True:
    try:
        button = chrome_browser.find_element_by_xpath('/html/body/nx-dialog/div[2]/div/taiwanlife-render-system-maintain-dialog/twlife-system-maintenance-modal/twlife-modal-frame/div[3]/div/div/rdr-button/button/div/div')
        ActionChains(chrome_browser).move_to_element(button).click(button).perform()
    except:
        break
        
while True:
    try:
        button = chrome_browser.find_element_by_xpath('/html/body/nx-dialog/div[2]/div/taiwanlife-render-system-maintain-dialog/twlife-system-maintenance-modal/twlife-modal-frame/div[3]/div/div/rdr-button/button/div/div')
        ActionChains(chrome_browser).move_to_element(button).click(button).perform()
    except:
        break
        
df_today = get_rate_type(yyyymm_today)


# 上月
chrome_browser.get(url)
time.sleep(2)

# 彈出視窗
while True:
    try:
        button = chrome_browser.find_element_by_xpath('/html/body/nx-dialog/div[2]/div/taiwanlife-render-system-maintain-dialog/twlife-system-maintenance-modal/twlife-modal-frame/div[3]/div/div/rdr-button/button/div/div')
        ActionChains(chrome_browser).move_to_element(button).click(button).perform()
    except:
        break
df_last = get_rate_type(yyyymm_last)

chrome_browser.close()


df_today = df_today[['商品名稱', '商品類別', str(yyyymm_today) +'宣告利率']]
df_last = df_last[['商品名稱', '商品類別', str(yyyymm_last) +'宣告利率']]

# check duplicates today
print(f'{yyyymm_today}宣告利率筆數：', len(df_today), sep='')
df_today1 = df_today.copy()
df_today1['counts'] = df_today1.groupby('商品名稱')['商品名稱'].transform('count')
df_today1.loc[df_today1['counts']>1]


# check duplicates last
print(f'{yyyymm_last}宣告利率筆數：', len(df_last), sep='')
df_last1 = df_last.copy()
df_last1['counts'] = df_last1.groupby('商品名稱')['商品名稱'].transform('count')
df_last1.loc[df_last1['counts']>1]


# check duplicates(有重覆再跑)
df_today_check = df_today1.loc[df_today1['counts']==1]
if len(df_today_check) != len(df_today):
    df_today2 = df_today.drop_duplicates(subset=['商品名稱'])
    print(f'{yyyymm_today}重複筆數：', len(df_today1.loc[df_today1['counts']!=1]), sep='') 
    print(f'{yyyymm_today}去重複筆數：', len(df_today2))  
else:
    df_today2 = df_today.copy()
    print(f'{yyyymm_today}無重複')
    

df_last_check = df_last1.loc[df_last1['counts']==1]
if len(df_last_check) != len(df_last):
    df_last2 = df_last.drop_duplicates(subset=['商品名稱']) 
    print(f'{yyyymm_last}重複筆數：', len(df_last1.loc[df_last1['counts']!=1]), sep='')
    print(f'{yyyymm_last}去重複筆數：', len(df_last2))  
else:
    df_last2 = df_last.copy()
    print(f'{yyyymm_last}無重複')
  
df_last2 = df_last2[['商品名稱', str(yyyymm_last) +'宣告利率']]
df1 = df_today2.merge(df_last2, how='left', on='商品名稱')
df1['壽險公司'] = company
df2 = df1[['壽險公司', '商品名稱', '商品類別', str(yyyymm_today) +'宣告利率', str(yyyymm_last) +'宣告利率']].copy()
df2



# 利率變動

df2['diff'] = df2[str(yyyymm_today) +'宣告利率'].str.replace('%', '').astype('float64') - df2[str(yyyymm_last) +'宣告利率'].str.replace('%', '').astype('float64')
df2['diff'].round(2)
df2.loc[df2['diff'] > 0, '升/降'] = '升'
df2.loc[df2['diff'] < 0, '升/降'] = '降'
df2.loc[df2['diff'] == 0, '升/降'] = ''
df2.loc[df2[str(yyyymm_last) +'宣告利率'].isna(), '升/降'] = '新商品/改版'

df2['幅度'] = abs(df2['diff'])*100
df2['幅度'] = df2['幅度'].loc[~df2['幅度'].isna()].round(decimals = 0).astype('int').astype('str') + 'bps'
# df2['幅度'].loc[df2['幅度'].isna()] = df2['升/降']
df2[['壽險公司', '商品名稱', '商品類別', str(yyyymm_today) +'宣告利率', str(yyyymm_last) +'宣告利率', '升/降', '幅度']]
df2


df3 = df2.copy()
df3['幣別'] = df3['商品類別'].str.split('（', expand=True)[1]
df3['幣別'] = df3['幣別'].str.strip('）')
df3['商品類別'] = df3['商品類別'].str.split('（', expand=True)[0]

df3 = df3[['壽險公司', '商品名稱', '商品類別', str(yyyymm_today) +'宣告利率', str(yyyymm_last) +'宣告利率', '升/降', '幅度']]
df3


df3['商品名稱'] = np.where(df3['商品名稱'].str.slice(0,4) == '台灣人壽', df3['商品名稱'].str.slice(4), df3['商品名稱'])
df3['商品名稱'] = df3['商品名稱'].str.replace('…', '|')
df3['商品名稱'] = df3['商品名稱'].str.replace('.', '|')

cond = df3['商品名稱'].str.contains('特定通路銷售')
df3['備註']     = np.where(cond, '特定通路銷售', None)
df3['商品名稱'] = np.where(cond, df3['商品名稱'].str.split('|', expand=True)[0], df3['商品名稱'])

df3['商品名稱'] = df3['商品名稱'].str.replace(' ', '')
df3['商品名稱'] = df3['商品名稱'].str.replace('fun心', 'Fun心')
df3['商品名稱'] = df3['商品名稱'].str.replace('ｅ富保', 'e富保')
df3['商品名稱'] = df3['商品名稱'].str.replace('年期', '年期')
df3['商品名稱'] = df3['商品名稱'].str.replace('醫療', '醫療')
df3

