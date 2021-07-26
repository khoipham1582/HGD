from selenium import  webdriver
from time import sleep
import random
from selenium.webdriver.common.keys import Keys
import pandas as pd
browser = webdriver.Chrome(executable_path="chromedriver.exe")
xl=pd.ExcelFile("huygt.xlsx")
df=pd.read_excel(xl,0,header=None)
#soluong=len(df.iloc[:,0]
# đăng nhập hgđ
browser.get("https://hogiadinh.baohiemxahoi.gov.vn/")
txtmail=browser.find_element_by_id("user_name")
#txtmail.send_keys("llthanh")
txtpass=browser.find_element_by_id("user_pass")
#txtpass.send_keys("123456a@A")
txtmacq=browser.find_element_by_id("id_ma_nh")
#txtmacq.send_keys("00008069")
sleep(30)
windows_before  = browser.current_window_handle
def huygt(mh,ms):
    cmnd=str(230280000+random.randint(1000,9999))
    phone="094366"+(str(random.randint(1000,9999)))
    # Mở trang quản lý HGĐ
    msg=""
    try:
        #obj = browser.switch_to.alert
        browser.get("https://hogiadinh.baohiemxahoi.gov.vn/InitQLNhanKhauHis.do")
        maho=browser.find_element_by_id("ma_ho_gd")
        maho.send_keys(str(mh))
        maho.send_keys(Keys.TAB)
        #obj.accept() 
        sleep(2)
        kbbd=browser.find_element_by_id("cmdAddNew").click()
        sleep(2)
        maso=browser.find_element_by_xpath("/html/body/section/article/form/aside[2]/div[1]/div[4]/input")
        maso.send_keys(str(ms))
        btntim=browser.find_element_by_id("cmdSearch")
        # mở trang hủy GT
        btntim.click()
        btnsua=browser.find_element_by_xpath("/html/body/section/article/form/aside[3]/table/tbody/tr[2]/td[3]/img")
        btnsua.click()
        #for winHandle in browser.window_handles:
        browser.switch_to.window(browser.window_handles[1])        
        socmnd=browser.find_element_by_id("so_cmtnd")
        browser.execute_script("arguments[0].value="+cmnd,socmnd)
        #socmnd.send_keys("230288266")
        tenchame=browser.find_element_by_id("ten_cha_me")
        #if tenchame.text!="": 
        browser.execute_script("arguments[0].value='bất kỳ'",tenchame)
        #tenchame.send_keys("bất kỳ")
        sdt=browser.find_element_by_id("phone_number")
        #if sdt.text!="": 
        arg="arguments[0].value='"+phone+"'"
        browser.execute_script(arg,sdt)
        #sleep(3)
        #sdt.send_keys("0943653668")
        #suatt=browser.find_element_by_id("cmdAddNew")
        browser.find_element_by_id("cmdKiemTra").click()    
        browser.find_element_by_id("cmdDeNghiDuyet").click()
        #obj.accept() 
        try:
            browser.find_element_by_id("cmdAddNew").click()
        except:
            print("")
        msg=browser.find_element_by_xpath("/html/body/div/table/tbody/tr/td/form/aside[2]/div[1]").text
        #msg=browser.find_element_by_class_name("msg-erorr").text 
        if "Trung CMND" in msg:
            if browser.current_window_handle != windows_before:
                browser.close()
                browser.switch_to.window(windows_before)
                huygt(mh,ms)
        if "Lỗi font chữ UNICODE" in msg:
            if browser.current_window_handle != windows_before:
                browser.close()
                browser.switch_to.window(windows_before)
                huygtdt(mh,ms)
        browser.close()  
        browser.switch_to.window(browser.window_handles[0])
        print("Mã số :",ms,msg)
    except:
        try:
            msg=browser.find_element_by_xpath("/html/body/div/table/tbody/tr/td/form/aside[2]/div[1]").text
            print("Hủy GT thất bại mã số ",ms,msg)
        except Exception as e: 
            #print('')
            if browser.current_window_handle != windows_before:
                browser.close()
                browser.switch_to.window(windows_before)
def huygtdt(mh,ms):
    cmnd=str(230280000+random.randint(1000,9999))
    phone="094366"+(str(random.randint(1000,9999)))
    # Mở trang quản lý HGĐ
    msg=""
    try:
        #obj = browser.switch_to.alert
        browser.get("https://hogiadinh.baohiemxahoi.gov.vn/InitQLNhanKhauHis.do")
        maho=browser.find_element_by_id("ma_ho_gd")
        maho.send_keys(str(mh))
        maho.send_keys(Keys.TAB)
        #obj.accept() 
        sleep(2)
        kbbd=browser.find_element_by_id("cmdAddNew").click()
        sleep(2)
        maso=browser.find_element_by_xpath("/html/body/section/article/form/aside[2]/div[1]/div[4]/input")
        maso.send_keys(str(ms))
        btntim=browser.find_element_by_id("cmdSearch")
        # mở trang hủy GT
        btntim.click()
        btnsua=browser.find_element_by_xpath("/html/body/section/article/form/aside[3]/table/tbody/tr[2]/td[3]/img")
        btnsua.click()
        #for winHandle in browser.window_handles:
        browser.switch_to.window(browser.window_handles[1])        
        socmnd=browser.find_element_by_id("so_cmtnd")
        browser.execute_script("arguments[0].value="+cmnd,socmnd)
        dantoc=browser.find_element_by_id("ma_dan_toc")
        browser.execute_script("arguments[0].value=10",dantoc)
        #socmnd.send_keys("230288266")
        tenchame=browser.find_element_by_id("ten_cha_me")
        #if tenchame.text!="": 
        browser.execute_script("arguments[0].value='bất kỳ'",tenchame)
        #tenchame.send_keys("bất kỳ")
        sdt=browser.find_element_by_id("phone_number")
        #if sdt.text!="": 
        arg="arguments[0].value='"+phone+"'"
        browser.execute_script(arg,sdt)
        #sleep(3)
        #sdt.send_keys("0943653668")
        #suatt=browser.find_element_by_id("cmdAddNew")
        browser.find_element_by_id("cmdKiemTra").click()    
        browser.find_element_by_id("cmdDeNghiDuyet").click()
        #obj.accept() 
        try:
            browser.find_element_by_id("cmdAddNew").click()
        except:
            print("")
        msg=browser.find_element_by_xpath("/html/body/div/table/tbody/tr/td/form/aside[2]/div[1]").text
        #msg=browser.find_element_by_class_name("msg-erorr").text 
        
        browser.close()  
        browser.switch_to.window(browser.window_handles[0])
        print("Mã số :",ms,msg)
    except:
        try:
            msg=browser.find_element_by_xpath("/html/body/div/table/tbody/tr/td/form/aside[2]/div[1]").text
            print("Hủy GT thất bại mã số ",ms,msg)
        except Exception as e: 
            #print('')
            if browser.current_window_handle != windows_before:
                browser.close()
                browser.switch_to.window(windows_before)
#print(df.head())
for i in range(len(df.iloc[:,0])):
    mh=str(df.at[i,0])
    if len(mh)==9:
        mh="0"+mh
    ms=str(df.at[i,1])
    if len(ms)==9:
        ms="0"+ms
    huygt(mh,ms)

