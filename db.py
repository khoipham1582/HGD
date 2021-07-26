import mysql.connector
mydb = mysql.connector.connect(
  host="localhost",
  user="root",
  password="",
  database="hgd"
)
mycursor=mydb.cursor()
sql="INSERT INTO `hgdtq` ( `masobhxh`,`hoten`, `ngaysinh`, `gioitinh`, `tks`, `hks`, `xks`, `thk`, `hhk`, `xhk`, `qh`, `lps`, `tt`, `nks`, `cmnd`, `ghichu`, `maho`, `stv`, `tench`) VALUES ( %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s,%s)"
val=('', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '','')
mycursor.execute(sql, val)
mydb.commit()
print(mycursor.rowcount, "record inserted.")
#print(mydb)