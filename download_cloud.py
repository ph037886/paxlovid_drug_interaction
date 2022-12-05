# -*- coding: utf-8 -*-
# python 3.8.10
#pip install selenium,pandas,requests,pywin32

from selenium import webdriver
from selenium.webdriver.support.ui import Select
import pandas as pd
from time import sleep
import pyautogui
import os

def download_cloud(password):
    driver= webdriver.Ie(r'D:/python/paxlovid/files/IEDriverServer')
    driver.get("https://medvpn.nhi.gov.tw/iwse0000/IWSE0001S01.aspx")
    pwd =str(password)
    card_in=driver.find_element("xpath",'//a[@href="IWSE0020S01.aspx"]')
    card_in.click()
    pwd_in=driver.find_element('id',"cph_ctl00_txtYSRYPIN")
    pwd_in.clear()
    pwd_in.send_keys(pwd)
    login=driver.find_element('id',"cph_ctl00_btnLogin")
    login.click()
    sleep(5)

    page =driver.get("https://medvpn.nhi.gov.tw/imme3000/IMME3000S01.aspx")
    close=driver.find_element('id',"btnClose3")
    close.click()
    df = pd.read_html(driver.page_source)
    cloud = df[8]
    driver.quit()
    nhi_med=pd.read_pickle(r'D:/python/paxlovid/files/nhi_med.pkl')[['藥品代號','藥品中文名稱','ATC_CODE']]

    cloud=cloud[['來源','成分名稱','藥品 健保代碼','藥品名稱','給藥 日數','藥品 用量','就醫(調劑)日期(住院用藥起日)']]
    cloud=cloud.join(nhi_med.set_index('藥品代號'), on='藥品 健保代碼') 
    cloud=cloud.rename(columns={'來源':'科別',
                                '成分名稱':'學名',
                                '藥品 健保代碼':'院內代碼',
                                '藥品名稱':'商品名',
                                '給藥 日數':'天數',
                                '藥品 用量':'總量',
                                '就醫(調劑)日期(住院用藥起日)':'日期',
                                '藥品中文名稱':'中文名'})
    cloud['醫師']="雲端藥歷"
    cloud[['身分證字號','姓名','每次劑量','途徑','頻率']]=''
    cloud=cloud[['身分證字號','姓名','日期','科別','醫師','院內代碼','商品名','學名','中文名','每次劑量','天數','總量','途徑','頻率','ATC_CODE']]
    return cloud

def next_page(password):
    try:
        driver= webdriver.Ie(executable_path=r'D:/python/paxlovid/files/IEDriverServer')
        driver.get("https://medvpn.nhi.gov.tw/iwse0000/IWSE0001S01.aspx")
    except:
        #IE的driver如果內容放大縮小，會跳error，這邊用鍵盤控制強制大小恢復100%，再關掉本來的，從新開一個新的
        pyautogui.hotkey('ctrl', '0')
        pyautogui.hotkey('alt', 'f4')
        print(1)
        driver= webdriver.Ie(executable_path=r'D:/python/paxlovid/files/IEDriverServer')
        driver.get("https://medvpn.nhi.gov.tw/iwse0000/IWSE0001S01.aspx")
        
    card_in=driver.find_element("xpath",'//a[@href="IWSE0020S01.aspx"]')
    card_in.click()
    pwd_in=driver.find_element('id',"cph_ctl00_txtYSRYPIN")
    pwd_in.clear()
    pwd =str(password)
    pwd_in.send_keys(pwd)
    login=driver.find_element('id',"cph_ctl00_btnLogin")
    login.click()
    sleep(5)

    page =driver.get("https://medvpn.nhi.gov.tw/imme3000/IMME3000S01.aspx")
    close=driver.find_element('id',"btnClose3")
    close.click()
    #雲端藥歷的網頁裡面有基本個資，屬性是hidden，一樣可以爬出來
    pt_id=driver.find_element('id',"ContentPlaceHolder1_HidUserID").get_attribute("value")
    pt_name=driver.find_element('id',"ContentPlaceHolder1_HidUserName").get_attribute("value")
    df = pd.read_html(driver.page_source)
    cloud = df[8]   
    try:
        select_var = list()
        for op in Select(driver.find_element('id','ctl00$ContentPlaceHolder1$pg_gvList_input')).options:
            select_var.append(op.text)
        for op in select_var:
            sleep(3)
            select = Select(driver.find_element('id','ctl00$ContentPlaceHolder1$pg_gvList_input'))
            select.select_by_visible_text(op)
            sleep(5)
            df = pd.read_html(driver.page_source)
            cloud=pd.concat([cloud,df[8]])
    except:
        pass
    driver.quit()
    nhi_med=pd.read_pickle(r'D:/python/paxlovid/files/nhi_med.pkl')[['藥品代號','藥品中文名稱','ATC_CODE']]

    cloud=cloud[['來源','成分名稱','藥品 健保代碼','藥品名稱','給藥 日數','藥品 用量','就醫(調劑)日期(住院用藥起日)']]
    cloud=cloud.join(nhi_med.set_index('藥品代號'), on='藥品 健保代碼') 
    cloud=cloud.rename(columns={'來源':'科別',
                                '成分名稱':'學名',
                                '藥品 健保代碼':'院內代碼',
                                '藥品名稱':'商品名',
                                '給藥 日數':'天數',
                                '藥品 用量':'總量',
                                '就醫(調劑)日期(住院用藥起日)':'日期',
                                '藥品中文名稱':'中文名'})
    cloud['醫師']="雲端藥歷"
    cloud['身分證字號']=pt_id
    cloud['姓名']=pt_name
    cloud[['每次劑量','途徑','頻率']]=''
    cloud=cloud[['身分證字號','姓名','日期','科別','醫師','院內代碼','商品名','學名','中文名','每次劑量','天數','總量','途徑','頻率','ATC_CODE']]
    return cloud,pt_id
