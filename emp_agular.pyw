# -*- coding: utf-8 -*-
import Tkinter as tk
from Tkinter import *
import MySQLdb as mydb
import tkFont
import datetime
#设置编码
import sys
reload(sys)
sys.setdefaultencoding('utf8')

#设置输入数据
term_list=[]
temp_list=[]


#设置项目题头
top=tk.Tk()
top.title("人员转正")
fm1=Frame(top)
fm1.grid(row=0,column=0)
fm2=Frame(top)
fm2.grid(row=1,column=0)
fm3=Frame(top)
fm3.grid(row=2,column=0)
fm4=Frame(top)
fm4.grid(row=3,column=0)


#设置字体
ft1 = tkFont.Font(family='msyhbd', size=15, weight=tkFont.BOLD)
ft2 = tkFont.Font(family='msyh', size=15)
ft3 = tkFont.Font(family='msyh', size=12)
ft4 = tkFont.Font(family='msyh', size=8)
ft5 = tkFont.Font(family='msyhbd', size=15)
#连接数据库
db_user=""
db_passwd=""
db_host="cdb-pdylikw0.bj.tencentcdb.com"
db_base="eddrac_employees"
db_port=10005
#上面是设置数据库参数

global cur
db=mydb.connect(host=db_host,user=db_user,passwd=db_passwd,port=db_port,db=db_base,charset='utf8')
cur=db.cursor()
#上面是建立数据库连接，打开指针

#构建查询语句
sql="SELECT all_employees.工号,all_employees.姓名,department.部门,position.职位,all_employees.入职时间 from all_employees,department,position where all_employees.部门=department.id_d and all_employees.职位=position.id_p and all_employees.转正日期 is NULL ORDER BY 入职时间 limit 12"

#执行查询
cur.execute(sql)

scan_result=cur.fetchall()

#列出查询结果
tk.Label(fm1,text="试用期人员列表",width=75,height=2,bg="#ffffff",font=ft1).grid(row=0,column=0,columnspan=6,sticky=N+S+E+W)
#空一行
tk.Label(fm2,text="",width=75,height=1,font=ft4).grid(row=0,column=0,columnspan=6,sticky=N+S+E+W)

#如果没有试用期员工
if len(scan_result)==0:
 tk.Label(top,text="没有试用期员工",width=60,height=5,bg="#aaaaaa",font=ft2).grid(row=1,column=0,columnspan=6,sticky=N+S+E+W)

#如果有试用期员工
else:
#列出题头
 Label(fm3,text="工号",font=ft3,width=6,relief=GROOVE).grid(row=0,column=0,sticky=N+S+E+W)
 Label(fm3,text="姓名",font=ft3,width=10,relief=GROOVE).grid(row=0,column=1,sticky=N+S+E+W)
 Label(fm3,text="部门",font=ft3,width=12,relief=GROOVE).grid(row=0,column=2,sticky=N+S+E+W)
 Label(fm3,text="职位",font=ft3,width=8,relief=GROOVE).grid(row=0,column=3,sticky=N+S+E+W)
 Label(fm3,text="入职时间",font=ft3,width=15,relief=GROOVE).grid(row=0,column=4,sticky=N+S+E+W)
 Label(fm3,text="转正日期",font=ft3,width=12,relief=GROOVE).grid(row=0,column=5,sticky=N+S+E+W)

 rl=1
 n=0
 while n<len(scan_result):
  Label(fm3,height=1).grid(row=rl,column=0,columnspan=6,sticky=N+S+E+W)
  Label(fm3,text=scan_result[n][0],font=ft3,width=6,relief=GROOVE).grid(row=rl+1,column=0,sticky=N+S+E+W)
  Label(fm3,text=scan_result[n][1],font=ft3,width=10,relief=GROOVE).grid(row=rl+1,column=1,sticky=N+S+E+W)
  Label(fm3,text=scan_result[n][2],font=ft3,width=12,relief=GROOVE).grid(row=rl+1,column=2,sticky=N+S+E+W)
  Label(fm3,text=scan_result[n][3],font=ft3,width=8,relief=GROOVE).grid(row=rl+1,column=3,sticky=N+S+E+W)
  Label(fm3,text=scan_result[n][4],font=ft3,width=15,relief=GROOVE).grid(row=rl+1,column=4)
  er=Entry(fm3,width=12,relief=GROOVE,font=ft3)
  er.grid(row=rl+1,column=5,sticky=E+W)
  temp_list.append(scan_result[n][0])
  temp_list.append(scan_result[n][1])
  temp_list.append(scan_result[n][4])
  temp_list.append(er)
  
  rl=rl+2
  n=n+1



#空行
Label(fm3,height=1,font=ft3,width=8).grid(row=rl,column=0,columnspan=6,sticky=N+S+E+W)


#日期检测函数
def date_chk(str_x):
 if len(str_x)==0:
  return 0
 if len(str_x)==10:
  if str_x[:4].isdigit() and str_x[5:7].isdigit() and str_x[-2:].isdigit():
    if int(str_x[:4])>2015 and int(str_x[:4]) < 2020 and int(str_x[5:7])<13 and int(str_x[5:7])>0 and  int(str_x[-2:])<32 and int(str_x[-2:])>0 and str_x[4:5]==str_x[-3:-2]=="-":
     return 1
 return -1 

#转正日期检测
def date_diff(x):
  x=x+" 00:00:00"
  y=str(datetime.datetime.now().date())+" 00:00:00"
  x=datetime.datetime.strptime(x,'%Y-%m-%d %H:%M:%S')
  y=datetime.datetime.strptime(y,'%Y-%m-%d %H:%M:%S')
  z=abs(x-y)
  if len(str(z))<15:
   return 0
  return int(str(z)[:2])




n_click=0


#确认提交函数
def ensure_auglar():
 global temp_list
 global en_but
 global n_click
 term_list
 global db
 
 n_chk=n_click%3
 r_text=""
 if n_chk==0:
  tl=0
  while tl<len(temp_list):
   au_date=temp_list[tl+3].get()
   if (date_chk(au_date)!=1 or date_diff(au_date)>7):
    
    temp_list[tl+3]=Entry(fm3,width=12,relief=GROOVE,fg="#ff0000",font=ft3)
    temp_list[tl+3].insert(END,"日期填写有误，暂不转正?")
    temp_list[tl+3].grid(row=tl/2+2,column=5)
   else:
    
    temp_list[tl+3]=Entry(fm3,width=12,relief=GROOVE,fg="#ff0000",font=ft3)
    temp_list[tl+3].insert(END,au_date)
    temp_list[tl+3].grid(row=tl/2+2,column=5)
   tl=tl+4
 elif n_chk==1:
  tl=0
  while tl<len(temp_list):
   au_date=temp_list[tl+3].get() 
   if (date_chk(au_date)!=1 or date_diff(au_date)>7):
    
    temp_list[tl+3]=Label(fm3,width=12,text="确认不予转正",relief=GROOVE,fg="#af0f30",font=ft3)
    
    temp_list[tl+3].grid(row=tl/2+2,column=5)
   else:
    
    temp_list[tl+3]=Entry(fm3,width=12,relief=GROOVE,fg="#00ff00",font=ft3)
    temp_list[tl+3].insert(END,au_date)
    temp_list[tl+3].grid(row=tl/2+2,column=5)
    term_list.append(temp_list[tl])
    term_list.append(au_date)
    term_list.append(temp_list[tl+1])
   tl=tl+4
  
 elif n_chk==2:
  tl=0
  r_text=""
  while 3*tl<len(term_list):
   
   sql="update all_employees set 转正日期=\""+term_list[3*tl+1]+"\" where 工号="+str(term_list[3*tl])
   
   cur.execute(sql)
   sql="insert into salary_mod(工号,确认时间,调整原因,之前基薪,之后基薪,之后绩效,之后全勤奖,之后岗位薪资,之后话补,之后补贴,之后车补)select 工号,转正时间,\"转正\",试用薪资,基薪,绩效,全勤奖,岗位工资,话补,补贴,车补 from all_employees where 工号="+str(term_list[3*tl])

   cur.execute(sql)
   r_text=r_text+term_list[3*tl+2]+","
   tl=tl+1
  db.commit() 
  r_text=r_text[:-1]
  if len(term_list)<1:
   Button(fm3,text="没有人转正",font=ft5,bg="#ff8000").grid(row=rl+1,column=0,columnspan=6,sticky=N+S+E+W)  
  else:
   Button(fm3,text=r_text+",转正成功",font=ft5,bg="#ff8000").grid(row=rl+1,column=0,columnspan=6,sticky=N+S+E+W)
 else:
  retrun
 n_click=n_click+1       



#确认按钮
#Label(fm3,text="",font=ft3,width=6).grid(row=rl,column=0,sticky=N+S+E+W)
#Label(fm3,text="",font=ft3,width=10).grid(row=rl,column=1,sticky=N+S+E+W)
#Label(fm3,text="",font=ft3,width=12).grid(row=rl,column=2,sticky=N+S+E+W)
#Label(fm3,text="",font=ft3,width=8).grid(row=rl,column=3,sticky=N+S+E+W)
#Label(fm3,text="",font=ft3,width=15).grid(row=rl,column=4,sticky=N+S+E+W)
en_but=Button(fm3,text="校验数据 准备转正",command=ensure_auglar,font=ft5,bg="#ff8000").grid(row=rl+1,column=0,columnspan=6,sticky=N+S+E+W)
#Label(fm3,text="",height=1,font=ft3,width=8).grid(row=rl+1,column=5)
Label(fm3,height=1).grid(row=rl+2,column=0,columnspan=6,sticky=N+S+E+W)





















top.mainloop()
