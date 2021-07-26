from selenium import  webdriver
from time import sleep
import random
from selenium.webdriver.common.keys import Keys
import pandas as pd
browser = webdriver.Chrome(executable_path="chromedriver.exe")
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException    
#from xlwt import Workbook
from openpyxl import Workbook
from datetime import datetime
#xl=pd.ExcelFile("ds2.xlsx")
#df=pd.read_excel(xl,0,header=None)
now=datetime.now()
#print(now.strftime("%d-%m-%H-%M-%S"))
k=1
j=99999999
wb = Workbook()
ws=wb.active
#ws = wb.add_sheet('ketqua') 
ws.write(0, 0, 'STT')
ws.write(0, 1, 'Mã số BHXH')
ws.write(0, 2, 'Họ tên')
ws.write(0, 3, 'Ngày sinh')
ws.write(0, 4, 'Giới tính')
ws.write(0, 5, 'Mã tỉnh KS')
ws.write(0, 6, 'Mã huyện KS')
ws.write(0, 7, 'Mã xã KS')
ws.write(0, 8, 'Mã tỉnh HK')
ws.write(0, 9, 'Mã huyện HK')
ws.write(0, 10, 'Mã xã HK')
ws.write(0, 11, 'QH chủ hộ')
ws.write(0, 12, 'Loại phát sinh')
ws.write(0, 13, 'Trạng thái')
ws.write(0, 14, 'Nơi khai sinh')
ws.write(0, 15, 'Số CMTND')
ws.write(0, 16, 'Ghi chú')
ws.write(0, 17, 'Mã hộ')
#soluong=len(df.iloc[:,0]
# đăng nhập hgđ
browser.get("https://tst.bhxh.gov.vn/")
#sleep(2)
#browser.find_element_by_id("bhxhMa").send_keys("06401")
browser.find_element_by_id("username").send_keys("cntt@gialai.vss.gov.vn")
browser.find_element_by_id("password").send_keys("H0G1@D1nh")
browser.find_element_by_id("password").send_keys(Keys.ENTER)
#browser.find_element_by_id("password").send_keys(Keys.F5)
#print("Bạn có 40s để đăng nhập.. !")
#sleep(40)
browser.get("https://tst.baohiemxahoi.gov.vn/#/qlNhanKhau")
#browser.find_element_by_xpath("/html/body/div[3]/div[1]/div/form/div[2]/div[2]/input").get_attribute('checked')
#tctq=browser.find_elements_by_css_selector("body > div.container > div.ng-scope > div > form > div.row.padding-10-top > div.col-sm-6.text-right > input")
#tctq=browser.find_elements_by_class_name("ng-pristine ng-valid ng-touched")
#tctq[1].click()
#tctq[1].click()
#print(*tctq,sep=",")
#sleep(3)
#xl=pd.ExcelFile("ds2.xlsx")
#df=pd.read_excel(xl,0,header=None)
now= now.strftime("%d-%m-%H-%M-%S")  
#print(df.head())
#for i in range(len(df.iloc[:,0])):
def check_exists_by_xpath(xpath):
    try:
        browser.find_element_by_xpath(xpath)
    except NoSuchElementException:
        return False
    return True
def traHGD(maho): 
    sleep(1)
    global k
    mh=browser.find_element_by_xpath("/html/body/div[3]/div[1]/div[2]/div[2]/div/fieldset/div[3]/div/div[1]/div/div/input")
    mh.send_keys(maho)
    #browser.execute_script("arguments[0].value="+maho,mh)
    #print("Tra HGD :"+maho)
    browser.find_element_by_xpath("/html/body/div[3]/div[1]/div[2]/div[2]/div/div/div/button[2]/span[2]").click()
    sleep(1)
    table=browser.find_element_by_xpath("/html/body/div[3]/div[1]/div[2]/div[3]/div/div/table")
    tr=table.find_elements_by_tag_name("tr")
    if(len(tr)>1):
        for i in range(len(tr)-1):
            #print(i+1)
            maso=tr[i+1].find_elements_by_tag_name("td")[1].text
            ht=tr[i+1].find_elements_by_tag_name("td")[2].text
            ns=tr[i+1].find_elements_by_tag_name("td")[3].text
            gt=tr[i+1].find_elements_by_tag_name("td")[4].text
            tks=tr[i+1].find_elements_by_tag_name("td")[5].text
            hks=tr[i+1].find_elements_by_tag_name("td")[6].text
            xks=tr[i+1].find_elements_by_tag_name("td")[7].text
            thk=tr[i+1].find_elements_by_tag_name("td")[8].text
            hhk=tr[i+1].find_elements_by_tag_name("td")[9].text
            xhk=tr[i+1].find_elements_by_tag_name("td")[10].text
            qh=tr[i+1].find_elements_by_tag_name("td")[11].text
            lps=tr[i+1].find_elements_by_tag_name("td")[12].text
            tt=tr[i+1].find_elements_by_tag_name("td")[13].text
            nks=tr[i+1].find_elements_by_tag_name("td")[14].text
            cmnd=tr[i+1].find_elements_by_tag_name("td")[15].text
            ghichu=tr[i+1].find_elements_by_tag_name("td")[16].text
            ws.write(k, 0, i+1)
            ws.write(k, 1, maso)
            ws.write(k, 2, ht)
            ws.write(k, 3, ns)
            ws.write(k, 4, gt)
            ws.write(k, 5, tks)
            ws.write(k, 6, hks)
            ws.write(k, 7, xks)
            ws.write(k, 8, thk)
            ws.write(k, 9, hhk)
            ws.write(k, 10, xhk)
            ws.write(k, 11, qh)
            ws.write(k, 12, lps)
            ws.write(k, 13, tt)
            ws.write(k, 14, nks)
            ws.write(k, 15, cmnd)
            ws.write(k, 16, ghichu)
            ws.write(k, 17, maho)
            k=k+1
    else:
        ws.write(k, 17, maho)
        ws.write(k, 16, "Không có nhân khẩu")
        k=k+1
    browser.execute_script("arguments[0].value=''",mh)
    #browser.find_element_by_xpath("/html/body/div[3]/div[1]/div[2]/div[2]/div/fieldset/div[3]/div/div[1]/div/div/input").send_keys("") 
while(j>99999980):
    #trahgd(i,df.at[i,0],df.at[i,1])
    #sleep(1)
    traHGD('01'+str(j))
    j=j-1
    #print('01'+i)
wb.save('HGD'+now+'.xlsx')
browser.close()