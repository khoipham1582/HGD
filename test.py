import random
import pandas as pd
phone="098665"+str(random.randint(1000,9999))
print("so dien thoai "+phone)
arg="arguments[0].value='"+phone+"'"
print(arg)
xl=pd.ExcelFile("huygt.xlsx")
df=pd.read_excel(xl,0,header=None)
s="0"+str(df.at[0,0])
print(len(s))
