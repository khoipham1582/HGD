from selenium import  webdriver
from time import sleep
import random
from selenium.webdriver.common.keys import Keys
import pandas as pd
browser=""
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException    
#from xlwt import Workbook
from openpyxl import Workbook
from datetime import datetime
import mysql.connector
import hashlib
mydb = mysql.connector.connect(
  host="localhost",
  user="root",
  password="",
  database="ehou"
)
mycursor=mydb.cursor()
now=datetime.now()
k=2
'''
browser.get("https://tst.bhxh.gov.vn/")
browser.find_element_by_id("username").send_keys("khoipnv@gialai.vss.gov.vn")
browser.find_element_by_id("password").send_keys("123321a@A")
browser.find_element_by_id("password").send_keys(Keys.ENTER)
browser.get("https://tst.baohiemxahoi.gov.vn/#/qlNhanKhau")
'''
def login(lk):
    #sleep(3)
    global browser
    browser = webdriver.Chrome(executable_path="chromedriver.exe")
    browser.get(lk)
    sleep(2)
    #browser.find_element_by_id("username").send_keys("cntt@gialai.vss.gov.vn")
    #browser.find_element_by_id("password").send_keys("H0G1@D1nh")
    #browser.find_element_by_id("password").send_keys(Keys.ENTER)
    #browser.get("https://tst.baohiemxahoi.gov.vn/#/qlNhanKhau")
now= now.strftime("%d-%m-%H-%M-")  
print("Nhập vào link cần tra cứu :")
lk=input()
login(lk)
#print("Nhập vào mã tỉnh từ 01-99")
#mt=input()
#print("Nhập vào stt hộ từ 99999999")
#j=input()
#j=int(j)
lk=lk.split('/')
mh=lk[len(lk)-1]
def trach(cl):
    global lk
    j=0
    i=1
    #ch=browser.find_elements_by_css_selector("body > div.fgid_130 > div > div > div > div.fgid_154 > div.fgid_160 > div.fgid_166")
    ch=browser.find_elements_by_class_name(cl)
    #ch1=browser.find_element_by_class_name("col-md-12 ml-auto mr-auto")
    #ch=browser.find_elements_by_xpath("/html/body/div[2]/div/div[2]/div/div[1]/div[1]/div[1]")
    print("Số câu hỏi tìm thấy",len(ch))
    while j<len(ch):
        if("check_box" in ch[j].text):
            td=ch[j].text.split('\n')[0]
            sda=len(ch[j].text.split('\n'))
            da1=''
            da2=''
            da3=''
            da4=''
            da5=''
            dad=''
            td_md5=''
            #print("Câu ",j,td,sda)
            for k in range(1,sda):
                if(k==1):da1=ch[j].text.split('\n')[k]
                if(k==2):da2=ch[j].text.split('\n')[k]
                if(k==3):da3=ch[j].text.split('\n')[k]
                if(k==4):da4=ch[j].text.split('\n')[k]
                if(k==5):da5=ch[j].text.split('\n')[k]
                if("check_box" in ch[j].text.split('\n')[k]):dad=ch[j].text.split('\n')[k]
                #print("Đáp án ",k,ch[j].text.split('\n')[k])
            td_md5=td+dad
            #td_md5=hashlib.md5(b'td_md5').hexdigest()
            sql1="INSERT INTO `data1`(`monhoc`,`tieude`, `da1`, `da2`, `da3`, `da4`, `da5`, `dad`,`td_md5`) VALUES(%s,%s,%s,%s,%s,%s,%s,%s,%s)" 
            val1=(mh,td,da1,da2,da3,da4,da5,dad,td_md5)
            #val1=('','','','','','','')
            try:
                kq=mycursor.execute(sql1, val1)
                print("Thêm câu hỏi thành công : ",td,kq)  
                #print(val1)
            except Exception as e:
                print("Thêm câu hỏi lỗi : ",td,e)  
            mydb.commit()           
            i=i+1
        j=j+1
    print("Số câu hỏi có đáp án",i)
trach("card-pricing")
trach("col-md-12")
browser.close()