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
sheet1.write(0, 1, 'Số BHXH')
sheet1.write(0, 2, 'Tên Công ty')
sheet1.write(0, 3, 'Địa chỉ')
sheet1.write(0, 4, 'Mức lương')
sheet1.write(0, 5, 'Hệ số lương')
sheet1.write(0, 6, 'Chức vụ')
sheet1.write(0, 7, 'Thời gian BHXH')
sheet1.write(0, 8, 'Ngày dừng đóng')
sheet1.write(0, 9, 'Từ tháng')
sheet1.write(0, 10, 'Đến tháng')
sheet1.write(0, 11, 'Mã Cty')
sheet1.write(0, 12, 'Thời gian BHTN')
sheet1.write(0, 13, 'PA')
sheet1.write(0, 14, 'Mức lương BHTN')
#soluong=len(df.iloc[:,0]
# đăng nhập hgđ
browser.get("https://tst.bhxh.gov.vn/")
sleep(2)
#browser.find_element_by_id("bhxhMa").send_keys("06401")
#browser.find_element_by_id("username").send_keys("cntt@gialai.vss.gov.vn")
#browser.find_element_by_id("password").send_keys("H0G1@D1nh")
#browser.find_element_by_id("password").send_keys(Keys.ENTER)
#browser.find_element_by_id("password").send_keys(Keys.F5)
print("Bạn có 40s để đăng nhập.. !")
sleep(40)
browser.get("https://tst.bhxh.gov.vn/#/tongHopQts")
#browser.find_element_by_xpath("/html/body/div[3]/div[1]/div/form/div[2]/div[2]/input").get_attribute('checked')
#tctq=browser.find_elements_by_css_selector("body > div.container > div.ng-scope > div > form > div.row.padding-10-top > div.col-sm-6.text-right > input")
#tctq=browser.find_elements_by_class_name("ng-pristine ng-valid ng-touched")
#tctq[1].click()
#tctq[1].click()
#print(*tctq,sep=",")
sleep(3)
xl=pd.ExcelFile("ds2.xlsx")
df=pd.read_excel(xl,0,header=None)
now= now.strftime("%d-%m-%H-%M-%S")+"SL"+str(len(df.iloc[:,0]))  
#print(df.head())
#for i in range(len(df.iloc[:,0])):
def check_exists_by_xpath(xpath):
    try:
        browser.find_element_by_xpath(xpath)
    except NoSuchElementException:
        return False
    return True
def traqt(i,soso,cq): 
    try: 
        #browser.find_element_by_name("soCmnd").send_keys("")
        sobhxh=browser.find_element_by_id("field_soSoBhxh")
        browser.execute_script("arguments[0].value=''",sobhxh)  
        sobhxh.send_keys(soso)
        #sleep(1)
        cqbhxh=browser.find_element_by_id("field_maBhxh")
        browser.execute_script("arguments[0].value=''",cqbhxh)
        cqbhxh.send_keys(cq)
        cqbhxh.send_keys(Keys.TAB)
        browser.find_element_by_xpath("/html/body/div[3]/div[1]/div/form/div[3]/div[2]/button[3]").click()  
        sleep(1)
        #browser.find_element_by_xpath("/html/body/div[3]/div[1]/div/form/div[2]/div[3]/button[2]/span").click()
        #sleep(1)
        #td=browser.find_element_by_xpath("/html/body/div[3]/div[1]/div/div[3]/table/tbody/tr[1]/td[2]")
        table=browser.find_element_by_xpath("/html/body/div[3]/div[1]/div/div[4]/table")
        tr=table.find_elements_by_tag_name("tr")
        tgbhxh=browser.find_element_by_xpath("/html/body/div[3]/div[1]/div/div[6]/div/label").text
        tgbhtn=browser.find_element_by_xpath("/html/body/div[3]/div[1]/div/div[7]/div/label").text
        qt=len(tr)
        if(qt>2):
            print("Tìm thấy dòng :",i,qt,tr[qt-1].text)
            #print("QT có số dòng :",qt)
            tuthang=tr[1].find_elements_by_tag_name("td")[1].text
            denthang=tr[qt-1].find_elements_by_tag_name("td")[2].text
            chucdanh=tr[qt-1].find_elements_by_tag_name("td")[8].text
            ngaynd=tr[qt-1].find_elements_by_tag_name("td")[2].text
            madv=tr[qt-1].find_elements_by_tag_name("td")[0].text
            PA=tr[qt-1].find_elements_by_tag_name("td")[5].text
            if(chucdanh==""):
                chucdanh=tr[qt-2].find_elements_by_tag_name("td")[8].text                
                if(chucdanh=="" and qt>3):
                    chucdanh=tr[qt-3].find_elements_by_tag_name("td")[8].text
                    if(chucdanh=="" and qt>4):
                        chucdanh=tr[qt-4].find_elements_by_tag_name("td")[8].text
            ml=tr[qt-1].find_elements_by_tag_name("td")[9].text
            mltn=tr[qt-1].find_elements_by_tag_name("td")[10].text
            if(ml=="" ):
                ml=tr[qt-2].find_elements_by_tag_name("td")[9].text
                if(ml=="" and qt>3):
                    ml=tr[qt-3].find_elements_by_tag_name("td")[9].text
                    if(ml=="" and qt>4):
                        ml=tr[qt-4].find_elements_by_tag_name("td")[9].text
            hsl=tr[qt-1].find_elements_by_tag_name("td")[11].text
            if(hsl==""):
                hsl=tr[qt-2].find_elements_by_tag_name("td")[11].text
                if(hsl=="" and qt>3):
                    hsl=tr[qt-3].find_elements_by_tag_name("td")[11].text
            cty=tr[qt-1].find_elements_by_tag_name("td")[23].text
            if(cty==""):
                cty=tr[qt-2].find_elements_by_tag_name("td")[23].text
            noilv=tr[qt-1].find_elements_by_tag_name("td")[24].text
            if(noilv==""):
                noilv=tr[qt-2].find_elements_by_tag_name("td")[24].text
                if(noilv=="" and qt>3):
                    noilv=tr[qt-3].find_elements_by_tag_name("td")[24].text
                    if(noilv=="" and qt>4):
                        noilv=tr[qt-4].find_elements_by_tag_name("td")[24].text
            
            sheet1.write(i, 0, i)
            sheet1.write(i, 1, soso)
            sheet1.write(i, 2, cty)
            sheet1.write(i, 3, noilv)
            sheet1.write(i, 4, ml)
            sheet1.write(i, 5, hsl)
            sheet1.write(i, 6, chucdanh)
            sheet1.write(i, 7, tgbhxh)
            sheet1.write(i, 8, ngaynd)
            sheet1.write(i, 9, tuthang)
            sheet1.write(i, 10, denthang)
            sheet1.write(i, 11, madv)
            sheet1.write(i, 12, tgbhtn)
            sheet1.write(i, 13, PA)
            sheet1.write(i, 14, mltn)
        elif (qt==2):
            print("Tìm thấy dòng :",i,qt,tr[qt-1].text)
            #print("QT có số dòng :",qt)
            td0=tr[1].find_elements_by_tag_name("td")[0].text
            if("Không có dữ liệu" not in td0):
                tuthang=tr[1].find_elements_by_tag_name("td")[1].text
                denthang=tr[qt-1].find_elements_by_tag_name("td")[2].text
                chucdanh=tr[qt-1].find_elements_by_tag_name("td")[8].text
                ngaynd=tr[qt-1].find_elements_by_tag_name("td")[2].text
                madv=tr[qt-1].find_elements_by_tag_name("td")[0].text
                PA=tr[qt-1].find_elements_by_tag_name("td")[5].text
                ml=tr[qt-1].find_elements_by_tag_name("td")[9].text
                hsl=tr[qt-1].find_elements_by_tag_name("td")[11].text
                cty=tr[qt-1].find_elements_by_tag_name("td")[23].text
                noilv=tr[qt-1].find_elements_by_tag_name("td")[24].text
                sheet1.write(i, 0, i)
                sheet1.write(i, 1, soso)
                sheet1.write(i, 2, cty)
                sheet1.write(i, 3, noilv)
                sheet1.write(i, 4, ml)
                sheet1.write(i, 5, hsl)
                sheet1.write(i, 6, chucdanh)
                sheet1.write(i, 7, tgbhxh)
                sheet1.write(i, 8, ngaynd)
                sheet1.write(i, 9, tuthang)
                sheet1.write(i, 10, denthang)
                sheet1.write(i, 11, madv)
                sheet1.write(i, 12, tgbhtn)
                sheet1.write(i, 13, PA)
            else:
                print("Không có dữ liệu")
        else:
            print("Không tìm thấy dòng : ",i)
            sheet1.write(i, 0, i)
            sheet1.write(i, 1, soso)
            sheet1.write(i, 2, "0")
        #sheet1.write(i, 2, ht)
        #sheet1.write(i, 3, ns)
        #sheet1.write(i, 4, so)
        #sheet1.write(i, 5, mt)
        #sheet1.write(i, 6, cq)
    except:
        print("Tìm lỗi")
        wb.save('ketqua'+now+'.xls')
     
for i in range(1,len(df.iloc[:,0])):
    traqt(i,df.at[i,0],df.at[i,1])
    #sleep(1)
wb.save('ketqua'+now+'.xls')
browser.close()