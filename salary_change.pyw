# -*- coding: utf-8 -*-
import sys
reload(sys)
sys.setdefaultencoding('utf8')
from Tkinter import *
import xlwt
import Tkinter as tk
import MySQLdb as mydb
import datetime
import subprocess
import tkFont

top=tk.Tk()
top.title("人员薪资调整")
fm1=Frame(top,relief=GROOVE,borderwidth=1)
fm1.grid(row=0,column=0)
fm2=Frame(top,relief=GROOVE)
fm3=Frame(top,relief=GROOVE)

#设置字体
ft1 = tkFont.Font(family='msyhbd',size=25)
ft2 = tkFont.Font(family='msyh',size=6)
ft3 = tkFont.Font(family='msyh',size=12)
ft4 = tkFont.Font(family='msyh',size=8)
ft5 = tkFont.Font(family='msyhbd',size=15)
ft6 = tkFont.Font(family='msyh',size=9)
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
n_click=0


#日期监测函数
def date_chk(str_x):

 if len(str_x)==0:
  return -1
 if len(str_x)==10:
     if str_x[:4].isdigit() and str_x[5:7].isdigit() and str_x[-2:].isdigit():
         if int(str_x[:4])>2015 and int(str_x[:4]) < 2020 and int(str_x[5:7])<13 and int(str_x[5:7])>0 and int(str_x[-2:])>0 and int(str_x[-2:])<32 and str_x[4:5]==str_x[-3:-2]=="-":
             return int(str_x[-2:])
 return -1

#日期有效性监测
def date_diff(x):
  x=x+" 00:00:00"
  y=str(datetime.datetime.now().date())+" 00:00:00"
  x=datetime.datetime.strptime(x,'%Y-%m-%d %H:%M:%S')
  y=datetime.datetime.strptime(y,'%Y-%m-%d %H:%M:%S')
  z=x-y
  if len(str(z))<15:
   return 0
  return int(str(z)[:-13])

#日期的数据库监测
def date_my_chk(x,y):
    sql="select 工号 from salary_mod where 确认时间>date_sub(\""+str(x)+"\",interval 2 month) and 工号=\""+str(y)+"\" and 调整原因!=\"新入职\" and 调整原因!=\"首次录入\" and 调整原因!=\"转正\""
    cur.execute(sql)
    l=cur.fetchall()
    if len(l)>0:
        return 0
    else:
        return 1
    
#检查工资数据
def num_chk(x):
    
    if x=="":
        return 0
    if len(x)<2:
        return -1
    if x.isdigit():
        if (len(x)<5 or (len(x)==5 and int(x[:1])<3)) and int(x[-1:])==0:
            return 1
        else:
            return -1
    else:
        return -1
    

#调整薪资提交函数
def salary_import():
    
    error_l=[]
    global fm_er
    
    try:
	fm_er.destroy()
    except NameError,e:
	pass
    fm_er=Frame(fm2)
    fm_er.grid(row=11,column=0,columnspan=8)
    
    if date_chk(scan_r[4].get())<1 or date_chk(scan_r[5].get())!=1:
        error_l.append("日期格式错误!")
        
    else:
        if date_diff(scan_r[4].get())>0 or date_diff(scan_r[5].get())<=0 or date_diff(scan_r[5].get())>=25:
            error_l.append("签字日期还没到，或者生效日期已过，日期非法!")
          
        else:
            if date_my_chk(scan_r[5].get(),scan_r[0])==0:
                error_l.append("该员工调薪速度过快,判定生效日期非法")
                
    lss=14
    if num_chk(str(scan_r[lss].get()))<1:
        error_l.append("调整后基本薪资填写有误!")
    lss=15
    s_title=["绩效","全勤奖","岗位工资","话补","补贴","车补"]
    while lss<20:
        if num_chk(str(scan_r[lss].get()))<0:
            error_l.append("调整后"+s_title[lss-14]+"填写有误!")
        lss=lss+1
    if len(error_l)>0:
	
        n=0
        for lss in error_l:
            Label(fm_er,text=lss,font=ft3).grid(row=n,column=0,columnspan=8)
            n=n+1
        return
    sql="insert into salary_mod values(Null,"+str(scan_r[0])+",\""+str(scan_r[4].get())+"\",\""+str(scan_r[5].get())+"\",\""+str(scan_r[6].get())+"\","
    lss=7
    while lss<14:
        if scan_r[lss]=="":
            scan_r[lss]="NULL"
        sql=sql+str(scan_r[lss])+","
        lss=lss+1
    while lss<21:
        scan_r[lss]=scan_r[lss].get()
        if scan_r[lss]=="":
            scan_r[lss]="NULL"
        sql=sql+str(scan_r[lss])+","
        lss=lss+1
    
    sql=sql[:-1]+")"
    cur.execute(sql)
    db.commit()
    salary_open()
    
    
            
        
            
            
   
            
        
    
    
    

#调整页面打开函数
def salary_open():
    
    global fm2,scan_r,fm3
    l_er=[]

    try:
	fm2.destroy()
    except NameError,e:
	pass
    try:
	fm3.destroy()
    except NameError,e:
	pass
    try:
	fm4.destroy()
    except NameError,e:
	pass

    fm2=Frame(top,relief=GROOVE)
    fm2.grid(row=0,column=1,sticky=N)
    fm3=Frame(top,relief=GROOVE)
    fm3.grid(row=1,column=1,sticky=N+S+W+E)
        
    sql1="select all_employees.工号,all_employees.姓名,department.部门,position.职位,all_employees.基薪,all_employees.绩效,all_employees.全勤奖,all_employees.岗位工资,all_employees.话补,all_employees.补贴,all_employees.车补 from all_employees inner join department on all_employees.部门=department.id_d inner join position on position.id_p=all_employees.职位 where all_employees.姓名 like \"%"+str(e_name.get())+"%\" and all_employees.当前=\"在职\""
    cur.execute(sql1)
    scan_r=list(cur.fetchall())


    


    
    if len(scan_r)!=1:
        
        error_l=Label(fm2,text="姓名输入有误",font=ft5)
        error_l.grid(row=0,column=0)
       
    else:
       
        scan_r=list(scan_r[0])
        
        
        Label(fm2,text=scan_r[2]+" "+scan_r[3]+" "+scan_r[1]+" 薪资调整表",relief=GROOVE,font=ft1,borderwidth=1).grid(row=0,column=0,columnspan=8,sticky=N+S+W+E)
        Label(fm2,font=ft3,relief=RAISED,borderwidth=1).grid(row=1,column=0,sticky=N+S+E+W)

        
        n=1
        for s_title in ["基薪","绩效","全勤奖","岗位工资","话补","补贴","车补"]:
            Label(fm2,text=s_title,relief=RAISED,font=ft5,borderwidth=1,width=8).grid(row=1,column=n,sticky=N+S+E+W)
            n=n+1
        Label(fm2,text="调整之前",font=ft5,borderwidth=2,relief=RAISED).grid(row=2,column=0,sticky=N+S+W+E)
        Label(fm2,text="调整之后",font=ft5,borderwidth=2,relief=RAISED).grid(row=3,column=0,sticky=N+S+W+E)
       
        


        
        Label(fm2,font=ft2).grid(row=4,column=0,sticky=N+S+E+W)
        Label(fm2,text="签字日期",font=ft5,relief=RAISED,borderwidth=2).grid(row=5,column=0,sticky=N+S+E+W)
        er=Entry(fm2,font=ft3,width=8)
        er.grid(row=5,column=1,sticky=N+S+W+E,columnspan=2)
        scan_r.insert(4,er)

        Label(fm2,text="生效日期",font=ft5,relief=RAISED,borderwidth=2).grid(row=5,column=4,sticky=N+S+E+W)
        er=Entry(fm2,font=ft3,width=8)
        er.grid(row=5,column=5,sticky=N+S+W+E,columnspan=2)
        scan_r.insert(5,er)
        
        Label(fm2,font=ft2).grid(row=6,column=0,sticky=N+S+E+W)
        Label(fm2,text="调整原因",font=ft5,borderwidth=2,relief=RAISED).grid(row=7,column=0,sticky=N+S+W+E)
        er=Entry(fm2,font=ft3)
        er.grid(row=7,column=1,sticky=N+S+W+E,columnspan=6)
        scan_r.insert(6,er)
        
        n=1
        while n< 8:
            if type(scan_r[n+6])==NoneType:
                scan_r[n+6]=""
            Label(fm2,text=str(scan_r[n+6]),font=ft3,relief=GROOVE,borderwidth=1,width=8).grid(row=2,column=n,sticky=N+S+W+E)
            er=Entry(fm2,font=ft3,width=8)
            er.grid(row=3,column=n,sticky=N+S+W+E)
            scan_r.append(er)
            n=n+1
        
        Label(fm2,font=ft2).grid(row=8,column=0,sticky=N+S+E+W)
        
        Button(fm2,font=ft5,text="确认提交",command=salary_import,bg="#30a330").grid(row=9,column=2,columnspan=3,sticky=N+S+W+E)
        Label(fm2,font=ft2).grid(row=10,column=0,sticky=N+S+E+W)

    
        list_s=str(scan_r[0])    
        sql2="select * from salary_mod where 工号=\""+list_s+"\""
        cur.execute(sql2)
        re_s=cur.fetchall()
        if len(re_s)==0:
            Label(fm3,text="该员工无薪资调整记录",font=ft1,relief=GROOVE).grid(row=0,column=0,sticky=N+S+W+E)
        else:
            
            re_s=list(re_s)
            for ql in range(len(re_s)):
                re_s[ql]=list(re_s[ql])
                
            Label(fm3,relief=RAISED,text="该员工薪资调动记录如下",font=ft1,borderwidth=1,width=8).grid(row=0,column=0,columnspan=9,sticky=N+S+E+W)
            n=2
            for s_title in ["基薪","绩效","全勤奖","岗位工资","话补","补贴","车补"]:
                Label(fm3,text=s_title,relief=RAISED,font=ft3,borderwidth=1,width=8).grid(row=1,column=n,sticky=N+S+E+W)
                n=n+1
            Label(fm3,font=ft3,borderwidth=1,relief=RAISED,width=8).grid(row=1,column=1,sticky=N+S+W+E)            
            Label(fm3,text="调整日期",font=ft3,borderwidth=1,relief=RAISED,width=8).grid(row=1,column=0,sticky=N+S+W+E)
            Label(fm3,text="调整之前",font=ft3,borderwidth=1,relief=GROOVE,width=8).grid(row=2,column=1,sticky=N+S+W+E)
            Label(fm3,text="调整之后",font=ft3,borderwidth=1,relief=GROOVE,width=8).grid(row=3,column=1,sticky=N+S+W+E)
            Label(fm3,text="调整原因",font=ft3,borderwidth=1,relief=GROOVE,width=8).grid(row=4,column=1,sticky=N+S+W+E)
        
            lr=0
        
            while lr< len(re_s):
                Label(fm3,text=re_s[lr][3],font=ft3,relief=GROOVE,borderwidth=1).grid(row=3*lr+2,column=0,rowspan=3,sticky=N+S+W+E)
                rl=5
                while rl<12:
                    if type(re_s[lr][rl])==NoneType:
                        re_s[lr][rl]=""
                    if type(re_s[lr][rl+7])==NoneType:
                        re_s[lr][rl+7]=""
                    Label(fm3,text=re_s[lr][rl],font=ft3,relief=GROOVE,borderwidth=1,width=8).grid(row=3*lr+2,column=rl-3)
                    Label(fm3,text=re_s[lr][rl+7],font=ft3,relief=GROOVE,borderwidth=1,width=8).grid(row=3*lr+3,column=rl-3)
                    rl=rl+1
                Label(fm3,text=re_s[lr][4],font=ft3,relief=GROOVE,borderwidth=1).grid(row=3*lr+4,column=2,columnspan=8,sticky=N+S+W+E)
                lr=lr+1

                                                                   
        
#薪资调整页面打开函数结束   
  
    
    


#薪资调整查看函数
def salary_scan():
   xx=0


#输入要调整薪资人员姓名
Label(fm1,text="输入要调整薪资的人员姓名",font=ft3,height=2,width=30).grid(row=0,column=0,columnspan=2,sticky=N+S+W+E)
Label(fm1,font=ft2).grid(row=1,column=0)
Label(fm1,font=ft3,text="姓名").grid(row=2,column=0,sticky=N+S+W+E)
e_name=Entry(fm1,font=ft3,width=15)
e_name.grid(row=2,column=1)
Label(fm1,font=ft2).grid(row=3,column=0,columnspan=2)
Button(fm1,text="确认",command=salary_open,borderwidth=5,bg="#208040").grid(row=4,column=0,columnspan=2,sticky=N+S+W+E)
Label(fm1,font=ft2).grid(row=5,column=0,columnspan=2)
Label(fm1,text="查看以往调整记录",font=ft3).grid(row=6,column=0,columnspan=2,sticky=N+S+W+E)
Label(fm1,font=ft2).grid(row=7,column=0,columnspan=2)
Label(fm1,font=ft3,text="调整日期").grid(row=8,column=0)
e_date=Entry(fm1,font=ft3,width=15)
e_date.grid(row=8,column=1)
Label(fm1,font=ft2).grid(row=9,column=0,columnspan=2)
Button(fm1,text="查看",command=salary_scan,borderwidth=5,bg="#a03020").grid(row=10,column=0,columnspan=2,sticky=N+S+W+E)

top.mainloop()
#sql=




