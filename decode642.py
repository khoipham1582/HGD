import base64
import pandas as pd
from xlwt import Workbook
from datetime import datetime
import xml.etree.ElementTree as ET
ten64=base64.b64encode(bytes('Phạm Nguyễn Việt Khôi',encoding="utf8"))
ten=base64.b64decode(ten64)
#print(ten,ten64)
wb = Workbook()

now=datetime.now()
now= now.strftime("%d-%m-%H-%M-%S")
tree = ET.ElementTree(file='bv.XML')
def giaima(i,dl,sheet1):
    gm64=''
    try:
        #gm64=base64.b64decode(dl+ b'=' * (-len(dl) % 4))
        #gm64=bytes(gm64,encoding="utf8")
        if len(dl)%4:
            dl += '=' * (4 - len(dl) % 4)
        gm64=base64.b64decode(dl)
        #gm64 += "=" * ((4 - len(dl) % 4) % 4)
        #print(gm64.decode()+'_VI_TINH></CHI_TIET_THUOC>')
        xmlTree=ET.fromstring(gm64.decode())
        j=0
        for e in xmlTree.iter():
            sheet1.write(0, j, e.tag)
            j=j+1
        j=0
        for e in xmlTree.iter():
            sheet1.write(i, j, e.text)
            j=j+1
            #print(e.tag,e.text)      
    except Exception as e: 
        print(e)
        #gm64=base64.b64decode(dl+ b'=' * (-len(dl) % 4))
        #gm64=base64.b64decode(dl+ b'===')
        #gm64=base64.b64decode(dl)
        #print("Decode lỗi !",i)
        #wb.save('decode'+now+'.xls')
    #print(type(gm64))
    #print(type(gm64.decode()))
#xl=pd.ExcelFile("bv.xlsx")
#df=pd.read_excel(xl,0,header=None)

#root=tree.getroot()
i=1
for e in tree.iter(tag='NOIDUNGFILE'):
    #print(e.text)
    sheet1 = wb.add_sheet('xml'+str(i)) 
    giaima(i,e.text,sheet1)
    i=i+1
wb.save('decode-'+str(i)+now+'.xls')
#for i in range(1,len(df.iloc[:,0])):
    #giaima(i,df.at[i,4])
    #wb.save('decode'+now+'.xls')

