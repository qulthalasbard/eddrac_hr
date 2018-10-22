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

top=tk.Tk()
top.title("人员离职")

#设置字体
ft1 = tkFont.Font(family='msyhbd',size=25)
ft2 = tkFont.Font(family='msyh',size=6)
ft3 = tkFont.Font(family='msyh',size=12)
ft4 = tkFont.Font(family='msyh',size=8)
ft5 = tkFont.Font(family='msyhbd',size=15)

#设置项目题头

fm1=Frame(top,relief=GROOVE,borderwidth=1)

fm3=Frame(top,relief=GROOVE,borderwidth=1)
fm4=Frame(top,relief=GROOVE,borderwidth=1)
message=0
#消息存储变量
#top=Frame(top)
#top.grid(row=0,column=1)

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


#上面是数据库连接，获取指针,获取查询结果


#日期监测函数
def date_diff(x):
  if date_chk(x)==0:
   return 0
  x=str(x)
  x=x+" 00:00:00"
  y=str(datetime.datetime.now().date())+" 00:00:00"
  x=datetime.datetime.strptime(x,'%Y-%m-%d %H:%M:%S')
  y=datetime.datetime.strptime(y,'%Y-%m-%d %H:%M:%S')
  z=x-y
  if len(str(z))<15:
   return 0
  return int(str(z)[:-13])

#日期格式监测函数
def date_chk(str_x):
 str_x=str(str_x)
 if len(str_x)==10:
   if str_x[:4].isdigit() and str_x[5:7].isdigit() and str_x[-2:].isdigit():
     if int(str_x[:4])>2015 and int(str_x[:4]) < 2020 and int(str_x[5:7])<13 and int(str_x[5:7])>0 and  int(str_x[-2:])<32 and int(str_x[-2:])>0 and str_x[4:5]==str_x[-3:-2]=="-":
      return 1
 return 0

#姓名在数据库是否存在且是否重复
def name_chk(str_x):
 sql1="select 工号 from all_employees where 姓名=\""+str_x+"\""
 cur.execute(sql1)
 x=len(cur.fetchall())
 if x==0:
  return 0
 sql2="select emp_quit.`工号`,all_employees.姓名 from emp_quit inner join all_employees on emp_quit.工号=all_employees.工号 where all_employees.姓名=\""+str_x+"\" and emp_quit.`确认时间` is null"
 cur.execute(sql2)
 y=len(cur.fetchall())
 if y==1:
  return 0
 else:
  return 1


ur_click=0

#急辞工函数
def urgent():
 global ur_click,urg
 ur_click +=1

 urg.destroy()
 if ur_click%3==0:
  urg=Button(fm1,text="真的急辞工?",font=ft3,command=urgent,bg="#903020")
  urg.grid(row=8,column=0,sticky=E+W+N+S)
 elif ur_click%3==1:
  urg=Button(fm1,text="确认急辞工!",font=ft3,command=urgent,bg="#a02000")
  urg.grid(row=8,column=0,sticky=E+W+N+S)
 else:
  urg=Label(fm1,text="急辞工",font=ft3,bg="#ff0000")
  urg.grid(row=8,column=0,sticky=E+W+N+S)


in_click=0
#数据库添加函数
def quit_import():
 global in_click,ur_click,urg
 global name,a_date,e_date,l_name,la_date,le_date
 global name_en,apply_date,expire_date
 
 if "fm2" in dir():
  fm2.destroy()
 if "fm4" in dir():
  fm4.destroy()
 if message!=0:
  message.destroy() 
 fm3=Frame(top,relief=GROOVE,borderwidth=1)
 fm3.grid(row=0,column=1,rowspan=15,sticky=E+W+N+S)
 
 n_chk = in_click%3
 if n_chk==0:
#点击第一次
  name=name_en.get()
  a_date=apply_date.get()
  e_date=expire_date.get()
  
 
  if name_chk(name)==0 or date_chk(a_date)==0 or date_chk(e_date)==0 or e_date < a_date or date_diff(a_date) < -7 or date_diff(a_date)>0 or date_diff(e_date)<0 or (date_diff(e_date)-date_diff(a_date)<30 and ur_click%3!=2):
   l_er=Label(fm3,text="姓名或日期填写错误",font=ft3)
   l_er.grid(row=0,column=0,sticky=N+S+E+W)
   in_click=0  
   return
  in_click=1
  return  
#点击第二次
 elif n_chk==1:
  l_name=Label(fm1,text=name,font=ft3)
  la_date=Label(fm1,text=a_date,font=ft3)
  le_date=Label(fm1,text=e_date,font=ft3)

  name_en.grid_remove()
  apply_date.grid_remove()
  expire_date.grid_remove()


  l_name.grid(row=1,column=1)
  la_date.grid(row=3,column=1)
  le_date.grid(row=5,column=1)   
  l_er=Label(fm3,text=str(name)+"于"+str(a_date)+"日申请离职,将于"+str(e_date)+"日左右正式离职。",font=ft3)
  l_er.grid(row=0,column=0)
  in_click=2
  
  return
#点击第三次
 elif n_chk==2:
  sql="insert into emp_quit(工号,申请时间,到期时间)select 工号,"+"\""+str(a_date)+"\",\""+str(e_date)+"\" from all_employees WHERE 姓名=\""+str(name)+"\""
  
  cur.execute(sql)
  db.commit()
  name_en=Entry(fm1,font=ft3,width=15)
  apply_date=Entry(fm1,font=ft3,width=15)
  expire_date=Entry(fm1,font=ft3,width=15)
  
  name_en.grid(row=1,column=1,sticky=N+S+E+W)
  apply_date.grid(row=3,column=1,sticky=N+S+E+W)
  expire_date.grid(row=5,column=1,sticky=N+S+E+W)
   
  l_er=Label(fm3,text=str(name)+"的离职申请登记成功",font=ft3)
  l_er.grid(row=0,column=0)
  in_click=3
  if ur_click%3==2:
   ur_click=ur_click+1
  print ur_click
  urg.destroy()
  print "OK"
  urg=Button(fm1,text="确认急辞工",font=ft3,command=urgent,bg="#7f5030")                                                                                                                                                
  urg.grid(row=8,column=0,sticky=E+W+N+S)                                                                                                                                                           
  name=a_date=e_date=0                                                                                                                                                                  
                                                                                                                                                                  
  
#查看已离职人员
def scan_qed():
 
#global qed_name
#global qed_date

 if "fm4" in dir():
  fm4.destroy()
 if "fm2" in dir():
  fm2.destroy()
 if "fm3" in dir():
  fm3.destroy()
 if message!=0:
  message.destroy() 

 fm4=Frame(top,relief=GROOVE,borderwidth=1)
 fm4.grid(row=0,column=1,rowspan=15,sticky=E+W+N+S)

 quit_time=qed_date.get()
 quit_name=qed_name.get()
 quit_de=qed_de.get()
#下面是列表题头
 Label(fm4,text="已离职人员列表",font=ft1,relief=GROOVE).grid(row=0,column=0,columnspan=6,sticky=N+E+W+S)

 Label(fm4,text="姓名",font=ft3,width=12,relief=GROOVE).grid(row=1,column=0,sticky=N)

 Label(fm4,text="部门",font=ft3,width=12,relief=GROOVE).grid(row=1,column=1)

 Label(fm4,text="职位",font=ft3,width=12,relief=GROOVE).grid(row=1,column=2)

 Label(fm4,text="电话",font=ft3,width=15,relief=GROOVE).grid(row=1,column=3)

 Label(fm4,text="申请日期",font=ft3,width=15,relief=GROOVE).grid(row=1,column=4)

 Label(fm4,text="离职日期",font=ft3,width=12,relief=GROOVE).grid(row=1,column=5)

#下面是数据库查询

 sql="select all_employees.姓名,department.`部门`,position.职位,all_employees.电话,emp_quit.申请时间,emp_quit.`确认时间` from emp_quit inner JOIN all_employees on emp_quit.`工号`=all_employees.`工号` inner JOIN department on all_employees.`部门`=department.id_d inner join position on all_employees.职位=position.id_p where emp_quit.`确认时间` like \"%"+quit_time+"%\" and all_employees.姓名 like \"%"+quit_name+"%\" and department.`部门` like \"%"+quit_de+"%\" order by emp_quit.`确认时间`"
 
 cur.execute(sql)
 quit_result=cur.fetchall()
 
#下面是列出数据 
 
 ql=0
 while ql< len(quit_result):
  Label(fm4,font=ft2).grid(row=2*ql+2,column=0)
  Label(fm4,font=ft2).grid(row=2*ql+3,column=0)
  xx=quit_result[ql][4]
  yy=quit_result[ql][5]
  Label(fm4,text=quit_result[ql][0],font=ft3,width=12).grid(row=2*ql+2,column=0,sticky=N)

  Label(fm4,text=quit_result[ql][1],font=ft3,width=12).grid(row=2*ql+2,column=1)

  Label(fm4,text=quit_result[ql][2],font=ft3,width=12).grid(row=2*ql+2,column=2)

  Label(fm4,text=quit_result[ql][3],font=ft3,width=15).grid(row=2*ql+2,column=3)

  Label(fm4,text=xx,width=15,font=ft3).grid(row=2*ql+2,column=4)
 
  Label(fm4,text=yy,font=ft3,width=12).grid(row=2*ql+2,column=5) 
  Label(fm4,font=ft2).grid(row=2*ql+3,column=0)
  ql += 1
 





#离职确认点击次数统计按钮
n_click=0

#下面是确认离职日期检查提交函数
def quit_confirm():
 global cur
 global db
 global scan_result
 global q_but
 global n_click
 
 global top
 n_chk=n_click%3
 if n_chk==0:
  rl=0
 
  while rl < len(scan_result):
   
   q_date=scan_result[rl][7].get()
  
   if date_chk(q_date)<-7 or date_diff(q_date)>7 or date_diff(q_date)-date_diff(scan_result[rl][6])<0 or date_diff(q_date)>0:
   
    scan_result[rl][7].delete(END,0)
    scan_result[rl][7]=Entry(fm2,font=ft3,width=12,bg="#af8030")
    scan_result[rl][7].insert(END,"暂不办理?")
    scan_result[rl][7].grid(row=2*rl+2,column=6)
   rl=rl+1
 if n_chk==1:
  rl=0
  while rl < len(scan_result):
   q_date=scan_result[rl][7].get()
   
   scan_result[rl][7].delete(0,END)
   if date_chk(q_date)<-7 or date_diff(q_date)>7 or date_diff(q_date)-date_diff(scan_result[rl][6])<0 or date_diff(q_date)>0:
    scan_result[rl][7]=0 
    
    Label(fm2,font=ft3,width=12,text="暂不办理!").grid(row=2*rl+2,column=6)
   else:
    scan_result[rl][7]=q_date
    Label(fm2,font=ft3,width=12,text=q_date).grid(row=2*rl+2,column=6)
   rl=rl+1
 if n_chk==2:
  rl=0
  qname=""
  
  while rl < len(scan_result):
   if scan_result[rl][7]!=0:
    sql1="update emp_quit set 确认时间=\""+str(scan_result[rl][7])+"\" where 工号=\""+str(scan_result[rl][0])+"\""
    sql2="update all_employees set 当前=\"离职\",宿舍=NULL where 工号=\""+str(scan_result[rl][0])+"\""
    sql_s_do="select 之后宿舍 from do_mod where next=NULL and 工号=\""+str(scan_result[rl][0])+"\""
    cur.execute(sql_s_do)
    result_s_do=cur.fetchall()
    if len(result_s_do)!=0:
     sql_q_do="insert into do_mod(工号,调整时间,调整原因,之前宿舍) values(\""+str(scan_result[rl][0])+"\",\""+str(scan_result[rl][7])+"\",\"离职\",\""+result_s_do[0][0]+"\")"
     cur.execute(sql_q_do)
    cur.execute(sql1)
    cur.execute(sql2)
    qname=qname+scan_result[rl][1]+","
   rl=rl+1
  db.commit()
  fm2.destroy
  
  if len(qname)==0:
    Label(top,text="没有确认任何人正式离职，重新办理离职手续请关闭该窗口重新打开",font=ft5,relief=GROOVE).grid(row=0,column=1,rowspan=15,sticky=E+W+N+S)
  else:
    message=Label(top,text=qname[:-1]+"离职确认完毕!",font=ft5,relief=GROOVE)
    message.grid(row=0,column=1,rowspan=15,sticky=E+W+N+S)
  scan_result=[]
 n_click += 1 


#下面是确认离职页面函数  
def scan_prequit():
 global n_click
 n_click=0
 global scan_result
 sql="select emp_quit.工号,all_employees.姓名,department.`部门`,position.职位,all_employees.电话,emp_quit.申请时间,emp_quit.`到期时间` from emp_quit inner JOIN all_employees on emp_quit.`工号`=all_employees.`工号` inner JOIN department on all_employees.`部门`=department.id_d inner join position on all_employees.职位=position.id_p where emp_quit.`确认时间` is NULL order by emp_quit.`申请时间` LIMIT 15"
 cur.execute(sql)
 scan_result=cur.fetchall()
#把查询结果改造为list
 scan_result=list(scan_result)
 for ns in range(len(scan_result)):
  scan_result[ns]=list(scan_result[ns])

 if fm3!=0:
  fm3.destroy()
 if fm4!=0:
  fm4.destroy()
 if message!=0:
  message.destroy()
 fm2=Frame(top,relief=GROOVE,borderwidth=1)
 global fm2
 fm2.grid(row=0,column=1,rowspan=15,sticky=E+W+N+S)
#在右侧打开离职人员列表+确认离职按钮
 r_title=Label(fm2,text="申请离职人员列表",font=ft1,relief=GROOVE).grid(row=0,column=0,columnspan=7,sticky=N+E+W+S)
 Label(fm2,text="姓名",font=ft3,width=12,relief=GROOVE).grid(row=1,column=0,sticky=N)

 Label(fm2,text="部门",font=ft3,width=12,relief=GROOVE).grid(row=1,column=1)

 Label(fm2,text="职位",font=ft3,width=12,relief=GROOVE).grid(row=1,column=2)

 Label(fm2,text="电话",font=ft3,width=15,relief=GROOVE).grid(row=1,column=3)

 Label(fm2,text="申请日期",font=ft3,width=15,relief=GROOVE).grid(row=1,column=4)

 Label(fm2,text="到期日期",font=ft3,width=12,relief=GROOVE).grid(row=1,column=5)


 Label(fm2,text="离职确认",font=ft3,width=12,relief=GROOVE).grid(row=1,column=6)
 
#下面是申请离职的人员
 rl=0
 rn=2
 while rl< len(scan_result):
  Label(fm2,text=scan_result[rl][1],font=ft3).grid(row=rl+rn,column=0)
  Label(fm2,text=scan_result[rl][2],font=ft3).grid(row=rl+rn,column=1)
  Label(fm2,text=scan_result[rl][3],font=ft3).grid(row=rl+rn,column=2)
  Label(fm2,text=scan_result[rl][4],font=ft3).grid(row=rl+rn,column=3)
  Label(fm2,text=scan_result[rl][5],font=ft3).grid(row=rl+rn,column=4)
  Label(fm2,text=scan_result[rl][6],font=ft3).grid(row=rl+rn,column=5)
  qen=Entry(fm2,font=ft3,width=12)
  qen.grid(row=rl+rn,column=6)
  scan_result[rl].append(qen)
  Label(fm2,font=ft2).grid(row=rl+rn+1,column=0)
  rl=rl+1
  rn=rn+1
 Label(fm2,font=ft2).grid(row=rl+rn,column=0)

 q_but=Button(fm2,text="检查数据，提交离职",command=quit_confirm,width=30,height=2,bg="#ef2a0f")
 q_but.grid(row=rl+rn,column=0,columnspan=8)


#下面是右侧离职申请添加按钮
fm1.grid(row=1,column=0)
Label(top,text="添加申请离职人员",font=ft3,relief=GROOVE).grid(row=0,column=0,sticky=N+S+E+W)
Label(fm1,text="",font=ft2).grid(row=0,column=0,columnspan=2,sticky=N+S+E+W)
Label(fm1,text="姓名",font=ft3).grid(row=1,column=0,sticky=N+S+E+W)
Label(fm1,text="",font=ft2).grid(row=2,column=0,sticky=N+S+E+W)
Label(fm1,text="申请日期",font=ft3).grid(row=3,column=0,sticky=N+S+E+W)
Label(fm1,text="",font=ft2).grid(row=4,column=0,sticky=N+S+E+W)
Label(fm1,text="到期日期",font=ft3).grid(row=5,column=0,sticky=N+S+E+W)
Label(fm1,text="",font=ft2).grid(row=6,column=0,sticky=N+S+E+W)
Button(fm1,text="确认添加",font=ft3,command=quit_import,bg="#7f5030").grid(row=7,column=0,columnspan=2,sticky=E+W+N+S)

#急辞工按钮
urg=Button(fm1,text="确认急辞工",font=ft3,command=urgent,bg="#7f5030")
urg.grid(row=8,column=0,sticky=E+W+N+S)
#离职确认按钮
Button(fm1,text="离职确认",font=ft3,command=scan_prequit,bg="#aaaa88").grid(row=8,column=1,columnspan=1,sticky=E+W+N+S)
#输入框
Label(fm1,text="",font=ft2).grid(row=9,column=0,sticky=N+S+E+W)
name_en=Entry(fm1,font=ft3,width=15)
apply_date=Entry(fm1,font=ft3,width=15)
expire_date=Entry(fm1,font=ft3,width=15)

name_en.grid(row=1,column=1)
apply_date.grid(row=3,column=1)
expire_date.grid(row=5,column=1)

#下面是查看已经离职人员

Label(fm1,text="",font=ft2).grid(row=10,column=0,sticky=N+S+E+W)
Label(fm1,text="姓名",font=ft3).grid(row=11,column=0,sticky=N+S+E+W)
Label(fm1,text="",font=ft2).grid(row=12,column=0,sticky=N+S+E+W)
Label(fm1,text="离职日期",font=ft3).grid(row=13,column=0,sticky=N+S+E+W)
Label(fm1,text="",font=ft2).grid(row=14,column=0,sticky=N+S+E+W)
qed_name=Entry(fm1,font=ft3,width=15)
qed_date=Entry(fm1,font=ft3,width=15)
qed_de=Entry(fm1,font=ft3,width=15)
qed_name.grid(row=11,column=1)
qed_date.grid(row=13,column=1)
qed_de.grid(row=15,column=1)
Label(fm1,text="部门",font=ft3).grid(row=15,column=0,sticky=N+S+E+W)
Label(fm1,text="",font=ft2).grid(row=16,column=0,sticky=N+S+E+W)
Button(fm1,text="查看已离职人员",font=ft3,command=scan_qed,width=28,bg="#1fa030").grid(row=17,column=0,columnspan=2,sticky=N+S+E+W)






#下面是查看申请离职人员



#下面是已经离职人员查看




    
  
      

   
 








top.mainloop()
