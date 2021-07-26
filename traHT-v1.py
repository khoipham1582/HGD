from selenium import  webdriver
from time import sleep
import random
from selenium.webdriver.common.keys import Keys
import pandas as pd
browser = ""
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException    
from xlwt import Workbook
from datetime import datetime
#xl=pd.ExcelFile("ds2.xlsx")
#df=pd.read_excel(xl,0,header=None)
now=datetime.now()
#print(now.strftime("%d-%m-%H-%M-%S"))

wb = Workbook()
sheet1 = wb.add_sheet('bhyt') 
sheet1.write(0, 0, 'STT')
sheet1.write(0, 1, 'Mã cơ quan')
sheet1.write(0, 2, 'Mã đơn vị')
sheet1.write(0, 3, 'Mã số BHXH')
sheet1.write(0, 4, 'Mã thẻ BHYT')
sheet1.write(0, 5, 'Họ tên')
sheet1.write(0, 6, 'Giới tính')
sheet1.write(0, 7, 'Ngày sinh')
sheet1.write(0, 8, 'Địa chỉ')
sheet1.write(0, 9, 'Từ ngày')
sheet1.write(0, 10, 'Đến ngày')
sheet1.write(0, 11, 'Mã tỉnh KCB')
sheet1.write(0, 12, 'Mã bệnh viện')
sheet1.write(0, 13, 'Ngày cấp')
sheet1.write(0, 14, 'Lần cấp')
def login():
    global browser
    browser=webdriver.Chrome(executable_path="chromedriver.exe")
    browser.get("https://tst.bhxh.gov.vn/")
    browser.find_element_by_id("username").send_keys("khoipnv@gialai.vss.gov.vn")
    browser.find_element_by_id("password").send_keys("123321a@A")
    browser.find_element_by_id("password").send_keys(Keys.ENTER)
    browser.get("https://tst.bhxh.gov.vn/#/ttCapThe")
    sleep(2)
login()
xl=pd.ExcelFile("dsht.xlsx")
df=pd.read_excel(xl,0,header=None)
now= now.strftime("%d-%m-%H-%M-%S")+"SL"+str(len(df.iloc[:,0]))
k=1
def traht(i,ht,ns):
    global k
    try:
        hoten=browser.find_element_by_xpath("/html/body/div[3]/div[1]/div/form/div[1]/div[4]/input")
        browser.execute_script("arguments[0].value=''",hoten)  
        hoten.send_keys(ht)
        ngaysinh=browser.find_element_by_xpath("/html/body/div[3]/div[1]/div/form/div[1]/div[5]/div/input")
        browser.execute_script("arguments[0].value=''",ngaysinh)  
        ngaysinh.send_keys(ns)
        browser.find_element_by_xpath("/html/body/div[3]/div[1]/div/form/div[2]/div[2]/button[2]").click()
        sleep(1)
        table=browser.find_element_by_xpath("/html/body/div[3]/div[1]/div/div[2]/table")
        tr=table.find_elements_by_tag_name("tr")
        sl=len(tr)
        if(len(tr)>1):
            print("Tìm thấy dòng :",i,tr[1].text)
            if("smsApp" in tr[1].text):
                browser.quit()
                login()
            #k=i*(len(tr))
            for j in range(1,len(tr)):
                macq=tr[j].find_elements_by_tag_name("td")[0].text
                madv=tr[j].find_elements_by_tag_name("td")[1].text
                masobhxh=tr[j].find_elements_by_tag_name("td")[2].text
                sothe=tr[j].find_elements_by_tag_name("td")[3].text
                gioitinh=tr[j].find_elements_by_tag_name("td")[5].text
                diachi=tr[j].find_elements_by_tag_name("td")[7].text
                tungay=tr[j].find_elements_by_tag_name("td")[8].text
                denngay=tr[j].find_elements_by_tag_name("td")[9].text
                tinhkcb=tr[j].find_elements_by_tag_name("td")[10].text
                mabv=tr[j].find_elements_by_tag_name("td")[11].text
                ngaycap=tr[j].find_elements_by_tag_name("td")[15].text
                lancap=tr[j].find_elements_by_tag_name("td")[13].text
                sheet1.write(k, 0, k)
                sheet1.write(k, 1, macq)
                sheet1.write(k, 2, madv)
                sheet1.write(k, 3, masobhxh)
                sheet1.write(k, 4, sothe)
                sheet1.write(k, 5, ht)
                sheet1.write(k, 6, gioitinh)
                sheet1.write(k, 7, ns)
                sheet1.write(k, 8, diachi)
                sheet1.write(k, 9, tungay)
                sheet1.write(k, 10, denngay)
                sheet1.write(k, 11, tinhkcb)
                sheet1.write(k, 12, mabv)
                sheet1.write(k, 13, ngaycap)
                sheet1.write(k, 14, lancap)
                k=k+1
            ts=browser.find_element_by_xpath("/html/body/div[3]/div[1]/div/div[3]/span").text
            ts=ts.replace("Tổng số : ","",1)
            ts=int(ts)
            if(ts>15):
                browser.find_element_by_xpath("/html/body/div[3]/div[1]/div/div[3]/ul/li[4]/a").click()
                sleep(1)
                tr=table.find_elements_by_tag_name("tr")
                for j in range(1,len(tr)):
                    macq=tr[j].find_elements_by_tag_name("td")[0].text
                    madv=tr[j].find_elements_by_tag_name("td")[1].text
                    masobhxh=tr[j].find_elements_by_tag_name("td")[2].text
                    sothe=tr[j].find_elements_by_tag_name("td")[3].text
                    gioitinh=tr[j].find_elements_by_tag_name("td")[5].text
                    diachi=tr[j].find_elements_by_tag_name("td")[7].text
                    tungay=tr[j].find_elements_by_tag_name("td")[8].text
                    denngay=tr[j].find_elements_by_tag_name("td")[9].text
                    tinhkcb=tr[j].find_elements_by_tag_name("td")[10].text
                    mabv=tr[j].find_elements_by_tag_name("td")[11].text
                    ngaycap=tr[j].find_elements_by_tag_name("td")[15].text
                    lancap=tr[j].find_elements_by_tag_name("td")[13].text
                    sheet1.write(k, 0, k)
                    sheet1.write(k, 1, macq)
                    sheet1.write(k, 2, madv)
                    sheet1.write(k, 3, masobhxh)
                    sheet1.write(k, 4, sothe)
                    sheet1.write(k, 5, ht)
                    sheet1.write(k, 6, gioitinh)
                    sheet1.write(k, 7, ns)
                    sheet1.write(k, 8, diachi)
                    sheet1.write(k, 9, tungay)
                    sheet1.write(k, 10, denngay)
                    sheet1.write(k, 11, tinhkcb)
                    sheet1.write(k, 12, mabv)
                    sheet1.write(k, 13, ngaycap)
                    sheet1.write(k, 14, lancap)
                    k=k+1
            if(ts>30):
                browser.find_element_by_xpath("/html/body/div[3]/div[1]/div/div[3]/ul/li[4]/a").click()
                sleep(1)
                tr=table.find_elements_by_tag_name("tr")
                for j in range(1,len(tr)):
                    macq=tr[j].find_elements_by_tag_name("td")[0].text
                    madv=tr[j].find_elements_by_tag_name("td")[1].text
                    masobhxh=tr[j].find_elements_by_tag_name("td")[2].text
                    sothe=tr[j].find_elements_by_tag_name("td")[3].text
                    gioitinh=tr[j].find_elements_by_tag_name("td")[5].text
                    diachi=tr[j].find_elements_by_tag_name("td")[7].text
                    tungay=tr[j].find_elements_by_tag_name("td")[8].text
                    denngay=tr[j].find_elements_by_tag_name("td")[9].text
                    tinhkcb=tr[j].find_elements_by_tag_name("td")[10].text
                    mabv=tr[j].find_elements_by_tag_name("td")[11].text
                    ngaycap=tr[j].find_elements_by_tag_name("td")[15].text
                    lancap=tr[j].find_elements_by_tag_name("td")[13].text
                    sheet1.write(k, 0, k)
                    sheet1.write(k, 1, macq)
                    sheet1.write(k, 2, madv)
                    sheet1.write(k, 3, masobhxh)
                    sheet1.write(k, 4, sothe)
                    sheet1.write(k, 5, ht)
                    sheet1.write(k, 6, gioitinh)
                    sheet1.write(k, 7, ns)
                    sheet1.write(k, 8, diachi)
                    sheet1.write(k, 9, tungay)
                    sheet1.write(k, 10, denngay)
                    sheet1.write(k, 11, tinhkcb)
                    sheet1.write(k, 12, mabv)
                    sheet1.write(k, 13, ngaycap)
                    sheet1.write(k, 14, lancap)
                    k=k+1
        else:
            print("Không tim thấy thông tin :",i,ht,ns)
    except Exception as e:
        print("Tìm lỗi",e)
        #browser.quit()
        #login()
        wb.save('ketqua'+now+'.xls')
        if("Timed out receiving message from renderer" in str(e)) or "Failed to execute script" in str(e) or "no such element: Unable to locate element" in str(e):
            browser.quit()
            login()
            
for i in range(2,len(df.iloc[:,0])):
    traht(i,df.at[i,0],df.at[i,1])
    #sleep(1)
wb.save('ketqua-bhyt'+now+'.xls')
browser.close()