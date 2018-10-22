# -*- coding: utf-8 -*-
import xlrd
import sys
reload(sys)
sys.setdefaultencoding('utf8')
from Tkinter import *
import xlwt
import Tkinter as tk
import openpyxl
import MySQLdb as mydb
import datetime
import subprocess
global term_entry
term_entry=[]
global title_data
label_before_l=[]
#导出数据路径
#export_path="r'f:\python\data.xls'"
#当前文件路径
#n_path="c:\python27\python.exe F:\\是\\err\\"

#结果导出类
class e_label:
 def __init__(self,text,width,base):
  self.text=text
  self.width=width
  self.base=base

 def c_label(self):
  return tk.Label(self.base,text=self.text,width=self.width,relief=GROOVE,borderwidth=1)

 def oncopy(self):
  top.clipboard_clear()
  top.clipboard_append(self.text)

 def c_menu(self):
  m=Menu(self.c_label(),tearoff=0)
  m.add_command(label="复制",command=self.oncopy)
  m.add_separator
  return m
  
 def p_menu(self,x):
  self.c_menu().post(x.x_root,x.y_root)
  

er_label=0
global er_label
#声明全局变量用来存储输入数据
 

#下面是真实长度函数
def real_len(str_xls):
 l=0
 if type(str_xls)==long:
  l=len(str(str_xls))
 elif type(str_xls)==datetime.datetime or type(str_xls)==datetime.date:
  l=12
 elif type(str_xls)==unicode:
  n=0
  while n< len(str_xls):
   str_n=str_xls[n:n+1]
   if str_n >= u'\u4e00' and str_n<=u'\u9fa5':
    l=l+2
   else:
    l=l+1 
   n=n+1
 else:
  l=6
 if l<6:
  l=6
 return int(l*1.3)
#上面函数判断字符真实视觉长度

#下面是类型转换函数
def data_str(str_xls):
 if type(str_xls)==long:
  l=str(str_xls)
 elif type(str_xls)==datetime.date or type(str_xls)==datetime.datetime :
  l=str(str_xls)[:10]
 elif type(str_xls)==NoneType:
  l=""
 else:
  l=str_xls
 return l
#类型转换函数

#匹配函数
def match_term(get_in,file_term):
  if get_in in file_term:
   return 1   
  else:
   return 0
#匹配函数



db_user=""
db_passwd=""
db_host="cdb-pdylikw0.bj.tencentcdb.com"
db_base="eddrac_employees"
db_port=10005

top=tk.Tk()
top.title("人员管理")
i=0

db=mydb.connect(host=db_host,user=db_user,passwd=db_passwd,port=db_port,db=db_base,charset='utf8')
cur=db.cursor()
global cur
#链接数据库，设置指针
cur.execute("select * from all_employees where 工号 like 1")
emp_data=cur.fetchall()
#读取数据
cur.execute('desc all_employees')
title_data=cur.fetchall()
#获取行宽
term_len=[]
xy=0
while xy < len(emp_data[0]):
 term_len.append(real_len(emp_data[0][xy]))
 xy=xy+1

#select查询数据是2重tuple

 


#读取字段title
while i <len(title_data)-1:
 tk.Label(top,text =title_data[i][0]).grid(row=i,column=0)
 er = tk.Entry(top)
 term_entry.append(er)
 er.grid(row=i,column=1)
#print type(title_data[i][0]),title_data[i][0]
 i=i+1
u=0
while u <len(title_data)-1:
 Label(top,text=title_data[u][0],width=real_len(emp_data[0][u])).grid(row=0,column=u+2)
 u=u+1
#print type(term_entry)
#print type(term_entry[0])
#测试  


#匹配函数
def match_term(get_in,file_term):
  if get_in in file_term:
   return 1   
  else:
   return 0
#清屏函数
def label_cls():
 global label_before_l
 if len(label_before_l)>0:
   k=0
   while k<len(label_before_l):
    s=0
    while s<len(label_before_l[k]):
     label_before_l[k][s].destroy()
     s=s+1
    k=k+1
 label_before_l=[]
 global  er_label
 if er_label!=0:
  er_label.destroy()
 er_label=0
  

#查找函数
def scan_employees():
  scan_result=[]
  global label_before_l
  global title_data
  global term_len
  global er_label
  label_cls()
  global scan_result
  
#获取输入数据
  l=0
  global term_entry
  term_entry_l=[]
  while l < len(term_entry):
   term_entry_l.append(term_entry[l].get())
   l=l+1
#检查部门，职位输入是否合法
  sql_de="select * from department where "+title_data[6][0]+" like \"%"+term_entry_l[6]+"%\""
  sql_posi="select * from position where "+title_data[7][0]+" like \"%"+term_entry_l[7]+"%\""
  cur.execute(sql_de)
  sql_de=cur.fetchall()
  cur.execute(sql_posi)
  sql_posi=cur.fetchall()
  if len(sql_posi)*len(sql_de)==0:
   er_label=Label(top,text="部门或职位填写错误")
   global er_label
   er_label.grid(row=2,column=3)
   return
#  print type(sql_de[1][0])
#  print sql_de[1][0]
#  print len(sql_de)

#构造查询语句
  sql=sql="select "
  u=0
  while u<6:
   sql=sql+title_data[u][0]+","
   u=u+1
  sql=sql+"(select "+title_data[6][0]+" from department where id_d=all_employees."+title_data[6][0]+") as "+title_data[6][0]+",(select "
  sql=sql+title_data[7][0]+" from position where id_p=all_employees."+title_data[7][0]+") as "+title_data[7][0]+","
  u=u+2
  while u<len(term_entry_l):
   sql=sql+title_data[u][0]+","
   u=u+1
  sql=sql[:-1]+" from all_employees where 1"
  u=0
  while u<len(term_entry_l):
   if u!=6 and u!=7 and len(term_entry_l[u])>0:
    sql=sql+" and "+title_data[u][0]+" like \"%"+term_entry_l[u]+"%\""
   u=u+1
  
  if len(term_entry_l[6])>0:
   sql=sql+" and "+title_data[6][0]+" in(select id_d from department where "+title_data[6][0]+" like \"%"+term_entry_l[6]+"%\")"
  if len(term_entry_l[7])>0:
   sql=sql+" and "+title_data[7][0]+" in(select id_p from position where "+title_data[7][0]+" like \"%"+term_entry_l[7]+"%\")"
  
 
  

  

  cur.execute(sql)
  scan_result=cur.fetchall()
  
#输出展示查询结果
  v=0 
  while v< len(scan_result):
   w=0
   label_before=[]
   while w<len(scan_result[0]):

     
    
     vw_label=e_label(data_str(scan_result[v][w]),term_len[w],top).c_label()
    
     vw_label.bind("<Button-3>",e_label(data_str(scan_result[v][w]),term_len[w],top).p_menu)
     
     label_before.append(vw_label)
     vw_label.grid(row=v+1,column=w+2)
     w=w+1
   label_before_l.append(label_before)
   v=v+1

#结果导出函数
def result_export():
  
  global title_data
  global scan_result
  book=xlwt.Workbook()
  sheet=book.add_sheet("scan_result",cell_overwrite_ok=True)
  n=0
  
  while n< len(title_data):
   sheet.write(0,n,title_data[n][0])
   n=n+1

  n=0
  while n< len(scan_result):
   m=0
   while m< len(scan_result[0]):
    sheet.write(n+1,m,data_str(scan_result[n][m]))
    m=m+1
   n=n+1

  book.save(r'C:\Users\love\Desktop\data.xls')
#考勤函数
def emp_attendance():
  cmd ="c:\python27\python.exe e:\2345\attendance.py"
  subprocess.Popen(cmd,shell=True,stdout=subprocess.PIPE)
 
#以下为数据导入所需的检测函数
#重复检测函数
def repeat_num(x):
 n=0
 m=0
 l=0
 while n <len(x):
  if x[n:n+1]==x[n+1:n+2]:
   m=m+1
  else:
   if m>l:
    l=m
   m=0
  n=n+1
 return l
#姓名检测函数
def name_chk(str_x):
 if len(str_x)>1 and len(str_x)<6:
  n=0
  m=0
  while n< len(str_x):
   if str_x[n:n+1] >= u'\u4e00' and str_x[n:n+1]<=u'\u9fa5':
    m=m+1
   n=n+1
  if m==n==len(str_x):
   
   return 1
  else:
   return 0
 else:
   return 0
#性别检测函数
def sex_chk(x):
 if x.encode("utf-8") in ["男","女"]:
  return 1
 return 0
#电话检测函数
def phone_chk(x):
 phoneprefix=["133","149","153","173","177","180","181","189","199","130","131","132","145","155","156","166","171","175","176","185","186","134","135","136","137","138","139","147","150","151","152","157","158","159","172","178","182","183","184","187","188","198"]
 if x.isdigit() and len(x)==11 and repeat_num(x)<6 and x[:3] in phoneprefix:
  return 1
 return 0
#身份证号码检测函数
def id_num_chk(str_x):
 id_check=["1","0","X","9","8","7","6","5","4","3","2"]
 if len(str_x)==18 and str_x[:17].isdigit() and int(str_x[6:10])<2002 and int(str_x[6:10])>1958 and  int(str_x[10:12])<13 and int(str_x[12:14])<32 and str_x[-1:] in id_check:
  id=[7,9,10,5,8,4,2,1,6,3,7,9,10,5,8,4,2 ]
  n=0 
  xy=0
  while n<17:
   xy=id[n]*int(str_x[n:n+1])+xy
   n=n+1
  n=xy%11
  if id_check[n]==str_x[-1:]:
   return 1
 return 0
#日期检测
def date_chk(str_x):
 if len(str_x)==10:
   if str_x[:4].isdigit() and str_x[5:7].isdigit() and str_x[-2:].isdigit():
     if int(str_x[:4])>2015 and int(str_x[:4]) < 2020 and int(str_x[5:7])<13 and int(str_x[5:7])>0 and  int(str_x[-2:])<32 and int(str_x[-2:])>0 and str_x[4:5]==str_x[-3:-2]=="-":
      return 1
 return 0 
#部门职位检测
def d_chk(x):
  sql_de="select id_d from department where "+title_data[6][0]+"= \""+x+"\""
  cur.execute(sql_de)
  sql_de=cur.fetchall()
  
  if len(sql_de)>0:
   return sql_de[0][0]
  return 0
#职位检测
def p_chk(x):
  
  
  sql_p="select id_p from position where "+title_data[7][0]+"= \""+x+"\""
  cur.execute(sql_p)
  sql_de=cur.fetchall()
  if len(sql_de)>0:
   return sql_de[0][0]
  return 0
#薪资检测
def s_chk(x):
  if x.isdigit() and int(x[-2:])==0 and int(x)<20001:
   return 1
  if x=="":
   return 0
  return -1
#宿舍检查
def do_chk(x):
  if x.isdigit():
   if (int(x)>300 and int(x)<319)or(int(x)>400 and int(x)<420):
    return 1
  return 0
#住址检测
def add_chk(str_x):
 if len(str_x)>4:
  n=0
  m=0
  while n< len(str_x):
   if str_x[n:n+1] >= u'\u4e00' and str_x[n:n+1]<=u'\u9fa5':
    m=m+1
   n=n+1
  if n-m<4:
   return 1
  else:
   return 0
 else:
  return 0

#输入框刷新函数
def en_refresh():
 for l in range(len(term_entry)):
  term_entry[l].delete(0,END)

#输入函数
def emp_import():
 l=0
 
 
 global term_entry
 term_entry_l=[]
 while l < len(term_entry):
  term_entry_l.append(term_entry[l].get())
  l=l+1


   
 if name_chk(term_entry_l[1])==0:
   label_cls()
   er_label=Label(top,text="姓名填得不对！")
   global er_label
   er_label.grid(row=2,column=3)
   return
 if sex_chk(term_entry_l[2])==0:
   label_cls()
   er_label=Label(top,text="性别填得不对！")
   global er_label
   er_label.grid(row=2,column=3)
   return
 if phone_chk(term_entry_l[3])+phone_chk(term_entry_l[19])<2:
   label_cls()
   er_label=Label(top,text="电话与紧急电话至少有1个填得不对！")
   global er_label
   er_label.grid(row=2,column=3)
   return
 if date_chk(term_entry_l[4])==0:
   label_cls()
   er_label=Label(top,text="入职日期填得不对！")
   global er_label
   er_label.grid(row=2,column=3)
   return
 if id_num_chk(term_entry_l[5])==0:
   label_cls()
   er_label=Label(top,text="身份证号码填得不对！")
   global er_label
   er_label.grid(row=2,column=3)
   return 0

#监测重复数据
 re_ck="select * from all_employees where "+title_data[5][0]+"=\""+term_entry_l[5]+"\""
 cur.execute(re_ck)
 if len(cur.fetchall())>0:
  label_cls()
  er_label=Label(top,text="该员工数据已录入")
  global er_label
  er_label.grid(row=2,column=3)
  return

  
#部门检测并替换
 entry_d=d_chk(term_entry_l[6])
 if entry_d==0:
   label_cls()
   er_label=Label(top,text="部门填得不对！")
   global er_label
   er_label.grid(row=2,column=3)
   return
 
#职位检测并替换
 entry_p=p_chk(term_entry_l[7])
 if entry_p==0:
   label_cls()
   er_label=Label(top,text="职位填得不对！")
   global er_label
   er_label.grid(row=2,column=3)
   return

#以下为从基薪到试薪的判断 
 sn=8
 while sn<16:
  if s_chk(term_entry_l[sn])==-1:
   err_t=title_data[sn][0]+"可以不填，但不能乱填!"
   label_cls()
   er_label=Label(top,text=err_t)
   global er_label
   er_label.grid(row=2,column=3)
   return
  sn=sn+1
 if s_chk(term_entry_l[15])+s_chk(term_entry_l[8])<1:
  label_cls()
  er_label=Label(top,text="基本薪资与试用期薪资应至少填写一个!")
  global er_label
  er_label.grid(row=2,column=3)
  return
#宿舍判断
 if do_chk(term_entry_l[16])==0 and len(term_entry_l[16])!=0:
  label_cls()
  er_label=Label(top,text="宿舍可以不填，但不能乱填!")
  global er_label
  er_label.grid(row=2,column=3)
  return
#地址检测
 if add_chk(term_entry_l[17])==0 or add_chk(term_entry_l[18])==0:
  label_cls()
  er_label=Label(top,text="家庭住址或籍贯必须认真填写！")
  global er_label
  er_label.grid(row=2,column=3)
  return
#介绍人检测
 if name_chk(term_entry_l[20])==0 and len(term_entry_l[20])!=0:
  label_cls()
  er_label=Label(top,text="介绍人可以不填，但不能乱填！")
  global er_label
  er_label.grid(row=2,column=3)
  return

  

#改造数据，忽略工号与转正日期，并将当前状态调整为试用
 term_entry_l[0]=""
 term_entry_l[22]="在职"
 term_entry_l[21]=""
 term_entry_l.append("80")
 import_l=1
 label_cls()

#构造执行语句
 sql_in="INSERT INTO all_employees values(null,"
 while import_l<24:
  if len(term_entry_l[import_l])!=0 and import_l!=6 and import_l!=7:
   sql_in=sql_in + "\""+term_entry_l[import_l]+"\","
  elif import_l==6:
   sql_in=sql_in+str(entry_d)+","
  elif import_l==7: 
   sql_in=sql_in+str(entry_p)+","
  else:
   sql_in=sql_in+"null,"
  import_l=import_l+1
 sql_in=sql_in[:-1]+")"

# f=open("sql.txt","w")
# print>>f,sql_in
# f.close()
# return

#开始执行添加并输出结果
 cur.execute(sql_in)
#没有下一句，结果不保存
 
#cur.close()
#更新工资调整表
 sql_u="insert into salary_mod(工号,确认时间,调整原因,之后基薪)select 工号,入职时间,\"新入职\",试用薪资 from all_employees where 身份证号码=\""+term_entry_l[5]+"\""
 cur.execute(sql_u)
#更新部门调整表
 sql_dd_u="insert into dd_mod(工号,调整时间,调整原因,之后部门,之后职位)select 工号,入职时间,\"新入职\",部门,职位 from all_employess where 身份证号码=\""+term_entry_l[5]+"\""
 cur.execute(sql_dd_u)
 
#更新宿舍调整表
 if len(str(term_entry_l[16]))!=0:
  sql_do_u="insert into do_mod(工号,调整时间,调整原因,之后宿舍)select 工号,入职时间,\"新入职\",宿舍 from all_employess where 身份证号码=\""+term_entry_l[5]+"\""
  cur.execute(sql_do_u)
 
 db.commit()
 label_cls()
 er_label=Label(top,text="添加成功！")
 global er_label
 er_label.grid(row=2,column=3)

#转正函数
def emp_agular():
 cmd="c:\python27\python.exe D:\Eddrac_hr\emp_agular.pyw"
 
 subprocess.Popen(cmd,shell=True,stdout=subprocess.PIPE)

#离职函数
def emp_quit():
 cmd="c:\python27\python.exe D:\Eddrac_hr\emp_quit.pyw"
 
 subprocess.Popen(cmd,shell=True,stdout=subprocess.PIPE)

#薪资调整函数
def sa_modify():
 cmd="c:\python27\python.exe D:\\Eddrac_hr\\salary_change.pyw"
 subprocess.Popen(cmd,shell=True,stdout=subprocess.PIPE)

#岗位调整函数
def p_modify():
 cmd="c:\python27\python.exe D:\\Eddrac_hr\\dd_mod.pyw"
 subprocess.Popen(cmd,shell=True,stdout=subprocess.PIPE)


#电费录入
def el_import():
 qq=0
#人员一览

def all_employee():
  cmd="c:\python27\python.exe D:\\Eddrac_hr\\all_employees.pyw"
 
  subprocess.Popen(cmd,shell=True,stdout=subprocess.PIPE)

#查找按钮
tk.Button(top,text='查找'.decode('utf-8'),command=scan_employees).grid(row=i,column=0,columnspan=2,sticky=E+W)
#导出按钮
tk.Button(top,text='导出'.decode('utf-8'),command=result_export).grid(row=1+i,column=0,columnspan=2,sticky=E+W)
#添加按钮
tk.Button(top,text='人员添加'.decode('utf-8'),command=emp_import).grid(row=i+2,column=0,sticky=E+W)
#考勤按钮
tk.Button(top,text='考勤管理'.decode('utf-8'),command=emp_attendance,bg="#00ff00").grid(row=i+2,column=1,sticky=E+W)
#刷新输入框按钮
tk.Button(top,text='清空输入框'.decode('utf-8'),command=en_refresh).grid(row=i+3,column=0,columnspan=1,sticky=E+W)
#刷新输入框按钮
tk.Button(top,text='清空查找结果'.decode('utf-8'),command=label_cls).grid(row=i+3,column=1,columnspan=1,sticky=E+W)
#转正按钮
tk.Button(top,text='人员转正'.decode('utf-8'),command=emp_agular,width=8).grid(row=i+4,column=0,columnspan=1,sticky=E+W)
#离职按钮
tk.Button(top,text='人员离职'.decode('utf-8'),command=emp_quit).grid(row=i+4,column=1,columnspan=1,sticky=E+W)
#薪资按钮
tk.Button(top,text='薪资调整'.decode('utf-8'),command=sa_modify).grid(row=i+5,column=0,columnspan=1,sticky=E+W)
#岗位按钮
tk.Button(top,text='宿舍及岗位调整'.decode('utf-8'),command=p_modify).grid(row=i+5,column=1,columnspan=1,sticky=E+W)
#宿舍调整按钮

#电费录入
tk.Button(top,text="电费录入".decode('utf-8'),command=el_import,bg="#00ff00").grid(row=i+6,column=0,columnspan=2,sticky=E+W)
#电费录入
tk.Button(top,text="人员一览表".decode('utf-8'),command=all_employee).grid(row=i+7,column=0,columnspan=2,sticky=E+W)
top.mainloop()
