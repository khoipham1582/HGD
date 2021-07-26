from selenium import  webdriver
from time import sleep
import random
from selenium.webdriver.common.keys import Keys
import pandas as pd
browser = webdriver.Chrome(executable_path="chromedriver.exe")
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException    
from xlwt import Workbook
from datetime import datetime
#xl=pd.ExcelFile("ds2.xlsx")
#df=pd.read_excel(xl,0,header=None)
now=datetime.now()
#print(now.strftime("%d-%m-%H-%M-%S"))

wb = Workbook()
sheet1 = wb.add_sheet('ketqua') 
sheet1.write(0, 0, 'STT')

#soluong=len(df.iloc[:,0]
# đăng nhập hgđ


#browser.find_element_by_id("bhxhMa").send_keys("06401")
#browser.find_element_by_id("username").send_keys("support@gialai.vss.gov.vn")
#browser.find_element_by_id("password").send_keys("P@ssw0rdcnttgl#")
#browser.find_element_by_id("password").send_keys(Keys.ENTER)
#browser.find_element_by_id("password").send_keys(Keys.F5)
#sleep(1)

#browser.get("https://tst.bhxh.gov.vn/#/tongHopQts")
#browser.find_element_by_xpath("/html/body/div[3]/div[1]/div/form/div[2]/div[2]/input").get_attribute('checked')
#tctq=browser.find_elements_by_css_selector("body > div.container > div.ng-scope > div > form > div.row.padding-10-top > div.col-sm-6.text-right > input")
#tctq=browser.find_elements_by_class_name("ng-pristine ng-valid ng-touched")
#tctq[1].click()
#tctq[1].click()
#print(*tctq,sep=",")
#sleep(3)
xl=pd.ExcelFile("dsdk.xlsx")
df=pd.read_excel(xl,0,header=None)
now= now.strftime("%d-%m-%H-%M-%S")+"SL"+str(len(df.iloc[:,0]))  
#print(df.head())
#for i in range(len(df.iloc[:,0])):
browser.get("https://dichvucong.baohiemxahoi.gov.vn/#/dang-ky") 
browser.find_element_by_xpath("/html/body/app-root/app-portal/div/div/app-dang-ky/div[2]/mat-horizontal-stepper/div[2]/div[1]/form/div[2]/button").click()
def dkgd(i,hoten,maso,cmnd,lh,sdt): 
    try:
        
        #browser.find_element_by_name("soCmnd").send_keys("")
        browser.find_element_by_xpath("/html/body/app-root/app-portal/div/div/app-dang-ky/div[2]/mat-horizontal-stepper/div[2]/div[2]/app-dang-ky-ca-nhan/form/div/div/div/div[2]/div/div/span/bhxh-input[1]/div/div/div[2]/mat-form-field/div/div[1]/div/input").send_keys(hoten)
        browser.find_element_by_xpath("/html/body/app-root/app-portal/div/div/app-dang-ky/div[2]/mat-horizontal-stepper/div[2]/div[2]/app-dang-ky-ca-nhan/form/div/div/div/div[2]/div/div/span/bhxh-input[2]/div/div/div[2]/mat-form-field/div/div[1]/div/input").send_keys(maso)
        browser.find_element_by_xpath("/html/body/app-root/app-portal/div/div/app-dang-ky/div[2]/mat-horizontal-stepper/div[2]/div[2]/app-dang-ky-ca-nhan/form/div/div/div/div[2]/div/div/span/div/div/div[1]/mat-form-field/div/div[1]/div/input").send_keys(cmnd)
        browser.find_element_by_xpath("/html/body/app-root/app-portal/div/div/app-dang-ky/div[2]/mat-horizontal-stepper/div[2]/div[2]/app-dang-ky-ca-nhan/form/div/div/div/div[2]/div/div/bhxh-input[4]/div/div/div[2]/mat-form-field/div/div[1]/div/input").send_keys(lh)
        browser.find_element_by_xpath("/html/body/app-root/app-portal/div/div/app-dang-ky/div[2]/mat-horizontal-stepper/div[2]/div[2]/app-dang-ky-ca-nhan/form/div/div/div/div[2]/div/div/div[4]/div/div/mat-form-field/div/div[1]/div[1]/input").send_keys(sdt)
        print("Có 20s để hoàn thiện thông tin")
    except:
        print("Đăng ký lỗi !")
for i in range(1,len(df.iloc[:,0])):
    print("Người thứ ",i)
    dkgd(i,df.at[i,0],df.at[i,1],df.at[i,2],df.at[i,3],df.at[i,4])
    sleep(20)
#wb.save('ketqua'+now+'.xls')
browser.close()