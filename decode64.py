import base64
import pandas as pd
from xlwt import Workbook
from datetime import datetime
import xml.etree.ElementTree as ET
ten64=base64.b64encode(bytes('Phạm Nguyễn Việt Khôi',encoding="utf8"))
ten=base64.b64decode(ten64)
#print(ten,ten64)
wb = Workbook()
sheet1 = wb.add_sheet('ketqua') 
sheet1.write(0, 0, 'STT')
sheet1.write(0, 1, 'MA_LK')
sheet1.write(0, 2, 'HO_TEN')
sheet1.write(0, 3, 'NGAY_SINH')
sheet1.write(0, 4, 'GIOI_TINH')
sheet1.write(0, 5, 'DIA_CHI')
sheet1.write(0, 6, 'MA_THE')
sheet1.write(0, 7, 'MA_DKBD')
sheet1.write(0, 8, 'GT_THE_TU')
sheet1.write(0, 9, 'GT_THE_DEN')
sheet1.write(0, 10, 'TEN_BENH')
sheet1.write(0, 11, 'MA_BENH')
sheet1.write(0, 12, 'MA_BENHKHAC')
sheet1.write(0, 13, 'MA_LYDO_VVIEN')
sheet1.write(0, 14, 'MA_NOI_CHUYEN')
sheet1.write(0, 15, 'MA_TAI_NAN')
sheet1.write(0, 16, 'NGAY_VAO')
sheet1.write(0, 17, 'NGAY_RA')
sheet1.write(0, 18, 'SO_NGAY_DTRI')
sheet1.write(0, 19, 'KET_QUA_DTRI')
sheet1.write(0, 20, 'TINH_TRANG_RV')
sheet1.write(0, 21, 'NGAY_TTOAN')
sheet1.write(0, 22, 'T_THUOC')
sheet1.write(0, 23, 'T_VTYT')
sheet1.write(0, 24, 'T_TONGCHI')
sheet1.write(0, 25, 'T_BNTT')
sheet1.write(0, 26, 'T_BNCCT')
sheet1.write(0, 27, 'T_BHTT')
sheet1.write(0, 28, 'T_NGUONKHAC')
sheet1.write(0, 29, 'T_NGOAIDS')
sheet1.write(0, 30, 'NAM_QT')
sheet1.write(0, 31, 'THANG_QT')
sheet1.write(0, 32, 'MA_LOAI_KCB')
sheet1.write(0, 33, 'MA_KHOA')
sheet1.write(0, 34, 'MA_CSKCB')
sheet1.write(0, 35, 'MA_KHUVUC')
sheet1.write(0, 36, 'MA_PTTT_QT')
sheet1.write(0, 37, 'CAN_NANG')

def giaima(i,dl):
    gm64=''
    try:
        gm64=base64.b64decode(dl)
        #gm64=bytes(gm64,encoding="utf8")
        #print(gm64)
        xmlTree=ET.fromstring(gm64.decode())
        sheet1.write(i, 0, i)
        for e in xmlTree.findall("MA_LK"):
            sheet1.write(i, 1, e.text)
        for e in xmlTree.findall("HO_TEN"):
            sheet1.write(i, 2, e.text)
        for e in xmlTree.findall("NGAY_SINH"):
            sheet1.write(i, 3, e.text)
        for e in xmlTree.findall("GIOI_TINH"):
            sheet1.write(i, 4, e.text)
        for e in xmlTree.findall("DIA_CHI"):
            sheet1.write(i, 5, e.text)
        for e in xmlTree.findall("MA_THE"):
            sheet1.write(i, 6, e.text)
        for e in xmlTree.findall("MA_DKBD"):
            sheet1.write(i, 7, e.text)
        for e in xmlTree.findall("GT_THE_TU"):
            sheet1.write(i, 8, e.text)
        for e in xmlTree.findall("GT_THE_DEN"):
            sheet1.write(i, 9, e.text)
        for e in xmlTree.findall("TEN_BENH"):
            sheet1.write(i, 10, e.text)
        for e in xmlTree.findall("MA_BENH"):
            sheet1.write(i, 11, e.text)
        for e in xmlTree.findall("MA_BENHKHAC"):
            sheet1.write(i, 12, e.text)
        for e in xmlTree.findall("MA_LYDO_VVIEN"):
            sheet1.write(i, 13, e.text)
        for e in xmlTree.findall("MA_NOI_CHUYEN"):
            sheet1.write(i, 14, e.text)
        for e in xmlTree.findall("MA_TAI_NAN"):
            sheet1.write(i, 15, e.text)
        for e in xmlTree.findall("NGAY_VAO"):
            sheet1.write(i, 16, e.text)
        for e in xmlTree.findall("NGAY_RA"):
            sheet1.write(i, 17, e.text)
        for e in xmlTree.findall("SO_NGAY_DTRI"):
            sheet1.write(i, 18, e.text)
        for e in xmlTree.findall("KET_QUA_DTRI"):
            sheet1.write(i, 19, e.text)
        for e in xmlTree.findall("TINH_TRANG_RV"):
            sheet1.write(i, 20, e.text)
        for e in xmlTree.findall("NGAY_TTOAN"):
            sheet1.write(i, 21, e.text)
        for e in xmlTree.findall("T_THUOC"):
            sheet1.write(i, 22, e.text)
        for e in xmlTree.findall("T_VTYT"):
            sheet1.write(i, 23, e.text)
        for e in xmlTree.findall("T_TONGCHI"):
            sheet1.write(i, 24, e.text)
        for e in xmlTree.findall("T_BNTT"):
            sheet1.write(i, 25, e.text)
        for e in xmlTree.findall("T_BNCCT"):
            sheet1.write(i, 26, e.text)
        for e in xmlTree.findall("T_BHTT"):
            sheet1.write(i, 27, e.text)
        for e in xmlTree.findall("T_NGUONKHAC"):
            sheet1.write(i, 28, e.text)
        for e in xmlTree.findall("T_NGOAIDS"):
            sheet1.write(i, 29, e.text)
        for e in xmlTree.findall("NAM_QT"):
            sheet1.write(i, 30, e.text)
        for e in xmlTree.findall("THANG_QT"):
            sheet1.write(i, 31, e.text)
        for e in xmlTree.findall("MA_LOAI_KCB"):
            sheet1.write(i, 32, e.text)
        for e in xmlTree.findall("MA_KHOA"):
            sheet1.write(i, 33, e.text)
        for e in xmlTree.findall("MA_CSKCB"):
            sheet1.write(i, 34, e.text)
        for e in xmlTree.findall("MA_KHUVUC"):
            sheet1.write(i, 35, e.text)
        for e in xmlTree.findall("MA_PTTT_QT"):
            sheet1.write(i, 36, e.text)
        for e in xmlTree.findall("CAN_NANG"):
            sheet1.write(i, 37, e.text)
    except:
        print("Decode lỗi !",i,gm64)
        #wb.save('decode'+now+'.xls')
    #print(type(gm64))
    #print(type(gm64.decode()))
xl=pd.ExcelFile("bv.xlsx")
df=pd.read_excel(xl,0,header=None)
now=datetime.now()
now= now.strftime("%d-%m-%H-%M-%S")+"SL"+str(len(df.iloc[:,0]))  
for i in range(1,len(df.iloc[:,0])):
    giaima(i,df.at[i,4])
    wb.save('decode'+now+'.xls')

