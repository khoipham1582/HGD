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
mydb = mysql.connector.connect(
  host="localhost",
  user="root",
  password="",
  database="hgd"
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
def login():
    #sleep(3)
    global browser
    browser = webdriver.Chrome(executable_path="chromedriver.exe")
    browser.get("https://tst.bhxh.gov.vn/")
    sleep(2)
    browser.find_element_by_id("username").send_keys("cntt@gialai.vss.gov.vn")
    browser.find_element_by_id("password").send_keys("H0G1@D1nh")
    browser.find_element_by_id("password").send_keys(Keys.ENTER)
    browser.get("https://tst.baohiemxahoi.gov.vn/#/qlNhanKhau")
def loadh():
    browser.get("https://tst.baohiemxahoi.gov.vn/#/qlNhanKhau")
    sleep(1)
now= now.strftime("%d-%m-%H-%M-")  
login()
def check_exists_by_xpath(xpath):
    try:
        browser.find_element_by_xpath(xpath)
    except NoSuchElementException:
        return False
    return True
def traHGD(maho): 
    try:   
        #sleep(1)
        global k
        mh=browser.find_element_by_xpath("/html/body/div[3]/div[1]/div[2]/div[2]/div/fieldset/div[3]/div/div[1]/div/div/input")
        mh.send_keys(maho)
        browser.find_element_by_xpath("/html/body/div[3]/div[1]/div[2]/div[2]/div/div/div/button[2]/span[2]").click()
        sleep(0.5)
        table=browser.find_element_by_xpath("/html/body/div[3]/div[1]/div[2]/div[3]/div/div/table")
        tr=table.find_elements_by_tag_name("tr")
        if(len(tr)>1):
            tch=browser.find_element_by_xpath("/html/body/div[3]/div[1]/div[2]/div[2]/div/fieldset/div[3]/div/div[2]/div").text
            stv=browser.find_element_by_xpath("/html/body/div[3]/div[1]/div[2]/div[4]/div/div[1]/button/span[2]").text
            for i in range(len(tr)-1):
                #print(i+1)
                maso=tr[i+1].find_elements_by_tag_name("td")[3].text
                ht=tr[i+1].find_elements_by_tag_name("td")[4].text
                ns=tr[i+1].find_elements_by_tag_name("td")[5].text
                gt=tr[i+1].find_elements_by_tag_name("td")[6].text
                tks=tr[i+1].find_elements_by_tag_name("td")[7].text
                hks=tr[i+1].find_elements_by_tag_name("td")[8].text
                xks=tr[i+1].find_elements_by_tag_name("td")[9].text
                thk=tr[i+1].find_elements_by_tag_name("td")[10].text
                hhk=tr[i+1].find_elements_by_tag_name("td")[11].text
                xhk=tr[i+1].find_elements_by_tag_name("td")[12].text
                qh=tr[i+1].find_elements_by_tag_name("td")[13].text
                lps=tr[i+1].find_elements_by_tag_name("td")[14].text
                tt=tr[i+1].find_elements_by_tag_name("td")[15].text
                nks=tr[i+1].find_elements_by_tag_name("td")[16].text
                cmnd=tr[i+1].find_elements_by_tag_name("td")[17].text
                ghichu=tr[i+1].find_elements_by_tag_name("td")[18].text
                sql="INSERT INTO `hgdtq` (`masobhxh`, `hoten`, `ngaysinh`, `gioitinh`, `tks`, `hks`, `xks`, `thk`, `hhk`, `xhk`, `qh`, `lps`, `tt`, `nks`, `cmnd`, `ghichu`, `maho`, `stv`, `tench`) VALUES(%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s,%s)"
                val=(maso, ht, ns, gt, tks, hks, xks, thk, hhk, xhk, qh, lps, tt, nks, cmnd, ghichu, maho, stv,tch)
                try:
                    mycursor.execute(sql, val)
                    print("Thêm nk thành công : ",maho,maso)  
                except:
                    print("Thêm nk lỗi : ",maho,maso)  
                mydb.commit()
                             
        else:
            print("Hộ không có nhân khẩu ",maho)
            k=k+1
        browser.execute_script("arguments[0].value=''",mh)
    except Exception as e:
        print("Lỗi : ",e)
        #if("element is not attached to the page document" in str(e)): loadh()
        if("Timed out receiving message from renderer" in str(e)) or "Failed to execute script thn-hgd7" in str(e) or "no such element: Unable to locate element" in str(e):
            browser.quit()
            login()
print("Nhập vào mã tỉnh từ 01-99")
mt=input()
print("Nhập vào stt hộ từ 99999999")
j=input()
j=int(j)
while(j>96000000):
    traHGD(mt+str(j))
    j=j-1
browser.close()