from selenium import  webdriver
from time import sleep
import random
from selenium.webdriver.common.keys import Keys
import pandas as pd
browser = webdriver.Chrome(executable_path="chromedriver.exe")
# đăng nhập hgđ
browser.get("https://hogiadinh.baohiemxahoi.gov.vn/")
txtmail=browser.find_element_by_id("user_name")
#txtmail.send_keys("khoi")
txtpass=browser.find_element_by_id("user_pass")
#txtpass.send_keys("Hvcv0102@")
txtmacq=browser.find_element_by_id("id_ma_nh");
#txtmacq.send_keys("00008069");
sleep(30)
windows_before  = browser.current_window_handle
def huygt(mh,ms):
    # Mở trang quản lý HGĐ
    try:
        browser.get("https://hogiadinh.baohiemxahoi.gov.vn/InitQLNhanKhauHis.do")
        maho=browser.find_element_by_id("ma_ho_gd")
        maho.send_keys(str(mh))
        maho.send_keys(Keys.TAB)
        sleep(2)
        kbbd=browser.find_element_by_id("cmdAddNew")
        kbbd.click()
        sleep(2)
        maso=browser.find_element_by_xpath("/html/body/section/article/form/aside[2]/div[1]/div[4]/input");
        maso.send_keys(str(ms))
        btntim=browser.find_element_by_id("cmdSearch")
        # mở trang hủy GT
        btntim.click()
        btnsua=browser.find_element_by_xpath("/html/body/section/article/form/aside[3]/table/tbody/tr[2]/td[3]/img");
        btnsua.click()

        for winHandle in browser.window_handles:
            browser.switch_to.window(winHandle)
    
        socmnd=browser.find_element_by_id("so_cmtnd")
        browser.execute_script("arguments[0].value='230288266'",socmnd)
        #socmnd.send_keys("230288266")
        tenchame=browser.find_element_by_id("ten_cha_me")
        #if tenchame.text!="": 
        browser.execute_script("arguments[0].value='bất kỳ'",tenchame)
        #tenchame.send_keys("bất kỳ")
        sdt=browser.find_element_by_id("phone_number")
        #if sdt.text!="": 
        browser.execute_script("arguments[0].value='0943653668'",sdt)
        #sdt.send_keys("0943653668")
        suatt=browser.find_element_by_id("cmdAddNew")
        suatt.click()
        obj = browser.switch_to.alert
        obj.accept()
        msg=browser.find_element_by_xpath("/html/body/div/table/tbody/tr/td/form/aside[2]/div[1]")
        print("Mã số : ",ms,msg.text)
        browser.close()
        #for winHandle in browser.window_handles:
        browser.switch_to.window(windows_before)
    except: print("Hủy GT thất bại mã số ",ms)
def xoacmnd():
    try:
        btnsua=browser.find_element_by_xpath("/html/body/section/article/form/aside[3]/table/tbody/tr[2]/td[3]/img")
        btnsua.click()
        for winHandle in browser.window_handles:
            browser.switch_to.window(winHandle)
        socmnd=browser.find_element_by_id("so_cmtnd")
        browser.execute_script("arguments[0].value=''",socmnd)
        #socmnd.send_keys("230288266")
        tenchame=browser.find_element_by_id("ten_cha_me")
        #if tenchame.text!="": 
        browser.execute_script("arguments[0].value=''",tenchame)
        #tenchame.send_keys("bất kỳ")
        sdt=browser.find_element_by_id("phone_number")
        #if sdt.text!="": 
        browser.execute_script("arguments[0].value=''",sdt)
        #sdt.send_keys("0943653668")
        suatt=browser.find_element_by_id("cmdAddNew")
        suatt.click()
        obj = browser.switch_to.alert
        obj.accept()
        msg=browser.find_element_by_xpath("/html/body/div/table/tbody/tr/td/form/aside[2]/div[1]")
        #print("Mã số 6421916916: "+msg.text)
        browser.close()
        browser.switch_to.window(windows_before)
    except: print("Xóa CMND thất bại !")
#huygt()
#xoacmnd()
xl=pd.ExcelFile("huygt.xlsx")
df=pd.read_excel(xl,0,header=None)
#print(df.head())
for i in range(len(df.iloc[:,0])):
    #print(i,df.at[i,0],df.at[i,1])
    #print("Dòng thứ ",i)
    huygt(df.at[i,0],str(df.at[i,1]))
    xoacmnd()